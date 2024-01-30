Attribute VB_Name = "myFunctions"
Public Enum opt_fx_Sum_or_Count_on_Single_Field
    Sum_Values
    Count_Values
End Enum

Option Explicit
Function fx_Error_Handler(Optional str_SubName As String, Optional bol_BreakCode As Boolean, Optional bol_Detailederr_Desc As Boolean)

' Purpose: This function will let the user know that an error occured and then provide some details for me.
' Trigger: Called Function
' Updated: 3/2/2022
'
' Change Log:
'       12/27/2021: Added the header info, Use Example, and updated name to fx_Error_Handler
'       2/2/2022:   Moved the Error related variables to the function so they don't have to be passed
'                   Updated so that the workbook name gets passed as part of the Error Soruce, and rmeoved the strTempVer
'                   Created the CodeBreak related code, and split the Messagebox
'                   Created the bol_Detailederr_Desc related code
'       3/2/2022:   Added  err_Num = 1004
'
' -----------------------------------------------------------------------------------------------------------------------------------
'
' Use Example: _
'    Call fx_Error_Handler(str_SubName:="o_11_Create_Dynamic_Customer_Array", bol_BreakCode:=True)
'
' Legend:
'   TBD:
'
' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare ErrorHandler Variables
    Dim str_WbName As String
        str_WbName = ThisWorkbook.Name

    Dim err_Source As String
        err_Source = Err.Source
    
    Dim err_Num As Long
        err_Num = Err.Number
    
    Dim err_Desc As String
        err_Desc = Err.Description
        
    Dim intMsgBoxButton As Long
    
    'Declare Detailed ErrorHandler Variables
    
    Dim strDetailedErrHandlerMessage As String

' -------------------------------------
' Determine which Messagebox to display
' -------------------------------------

    If bol_BreakCode = True Then
        GoTo ErrorHandler_CodeBreak
    Else
        GoTo ErrorHandler_NoCodeBreak
    End If

' ---------------------------
' Output the error messagebox
' ---------------------------

ErrorHandler_NoCodeBreak:

' Use this message if I choose not to break the code and allow the user to continue on
intMsgBoxButton = MsgBox( _
    Title:="I am Error", _
    Buttons:=vbCritical, _
    Prompt:="Something went awry, try to hit OK and redo the last step. " _
    & "If that doesn't resolve it then reach out to James Rinaldi for a fix. " _
    & "This tool has a growth mindset, with each issue addressed we itterate to a better version." & Chr(10) & Chr(10) _
    & "Please take a screenshot of this message, and send it to James. " _
    & "Include a brief description of what you were doing when it occurred." & Chr(10) & Chr(10) _
    & "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" & Chr(10) _
    & "--------------------------------------------------------------------------------" & Chr(10) _
    & "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" & Chr(10) & Chr(10) _
    & "Error Source: " & str_WbName & " - " & err_Source & Chr(10) _
    & "Subroutine: " & str_SubName & Chr(10) _
    & "Error Desc.: #" & err_Num & " - " & err_Desc & Chr(10))

GoTo SecondStage

' -------------------------------------------------
' Output the error messagebox if Break Code is true
' -------------------------------------------------
    
ErrorHandler_CodeBreak:

' Use this message if I choose  to break the code and don't allow the user to continue on
intMsgBoxButton = MsgBox( _
    Title:="I am Error", _
    Buttons:=vbCritical + vbOKCancel, _
    Prompt:="Something went awry, please hit OK and redo the last step, or Cancel and see if you code continues to run. " _
    & "If that doesn't resolve it then reach out to James Rinaldi for a fix. " _
    & "This tool has a growth mindset, with each issue addressed we itterate to a better version." & Chr(10) & Chr(10) _
    & "Please take a screenshot of this message, and send it to James. " _
    & "Include a brief description of what you were doing when it occurred." & Chr(10) & Chr(10) _
    & "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" & Chr(10) _
    & "--------------------------------------------------------------------------------" & Chr(10) _
    & "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" & Chr(10) & Chr(10) _
    & "Error Source: " & str_WbName & " - " & err_Source & Chr(10) _
    & "Subroutine: " & str_SubName & Chr(10) _
    & "Error Desc.: #" & err_Num & " - " & err_Desc & Chr(10))
    
GoTo SecondStage
    
' ---------------------------------------------------------
' Finish the process, depending on what button was selected
' ---------------------------------------------------------
    
SecondStage:
            
    If bol_Detailederr_Desc = True Then
        GoTo DetailedErrorHandler
    End If
    
ThirdStage:
    
    If intMsgBoxButton = 1 Then 'User selected OK
        myPrivateMacros.DisableForEfficiencyOff
        End
    ElseIf intMsgBoxButton = 2 Then 'User selected Cancel
        Stop
    End If
    
' -----------------------------------------
' Provide more info on specific error types
' -----------------------------------------
    
DetailedErrorHandler:
    
    If err_Num = 91 Then
        strDetailedErrHandlerMessage = _
            "Your error is the result of an Object not being set." & Chr(10) & Chr(10) & _
            "This typically happens becuase you didn't qualify a range, check for any ws.range(cells(),cells()) situations"
    
    ElseIf err_Num = 1004 Then
        strDetailedErrHandlerMessage = _
            "Your error is the result of a range that can't be found." & Chr(10) & Chr(10) & _
            "This typically happens becuase you didn't qualify a range, or your referencing a Named Range that doesn't exist."
    
    End If
        
    intMsgBoxButton = MsgBox( _
        Title:="Just an FYI", _
        Buttons:=vbInformation, _
        Prompt:=strDetailedErrHandlerMessage)
        
    GoTo ThirdStage
    
End Function
Function fx_Record_ErrorLog(Optional str_ModuleName As String, Optional str_SubName As String)

' Purpose: To record an error in the Error Log.
' Trigger: Workbook Change
' Updated: 3/16/2022
'
' Change Log:
'       3/15/2022:  Initial Creation
'       3/16/2022:  Added the str_ModuleName variable
'                   Passed the project's path to the Error Log
'
' -----------------------------------------------------------------------------------------------------------------------------------
'
' Use Example: _
'    Call fx_Record_ErrorLog(str_ModuleName:="ThisWorkbook", str_SubName:="o_11_Create_Dynamic_Customer_Array")
'
' Legend:
'   TBD:

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare Strings
    
    Dim err_Source As String
        err_Source = Err.Source
    
    Dim err_Num As Long
        err_Num = Err.Number
    
    Dim err_Desc As String
        err_Desc = Err.Description

' If I receive an error creating the error log just capture what I can
' Needs to be below setting the error variables so the new error doesn't erase the details of the original
On Error Resume Next

    'Declare Worksheets

    Dim ws_ErrorLog As Worksheet
    Set ws_ErrorLog = ThisWorkbook.Sheets("ERROR LOG")

    'Declare Integers

    Dim int_CurRow As Long
        int_CurRow = ws_ErrorLog.Range("A:A").Find("").Row
        
    Dim int_LastCol As Long
        int_LastCol = ws_ErrorLog.Cells(1, ws_ErrorLog.Columns.Count).End(xlToLeft).Column
        
    'Declare Cell References
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(ws_ErrorLog.Range(ws_ErrorLog.Cells(1, 1), ws_ErrorLog.Cells(1, int_LastCol)))

    Dim col_ErrorOccured As Long
        col_ErrorOccured = fx_Create_Headers("Error Occurred", arry_Header)

    Dim col_Errorer As Long
        col_Errorer = fx_Create_Headers("Errorer", arry_Header)

    Dim col_SubRoutine As Long
        col_SubRoutine = fx_Create_Headers("Subroutine", arry_Header)

    Dim col_ErrorSource As Long
        col_ErrorSource = fx_Create_Headers("Error Source", arry_Header)

    Dim col_ErrorNum As Long
        col_ErrorNum = fx_Create_Headers("Error #", arry_Header)

    Dim col_ErrorDesc As Long
        col_ErrorDesc = fx_Create_Headers("Error Desc.", arry_Header)
        
    Dim col_ProjectPath As Long
        col_ProjectPath = fx_Create_Headers("Project Path", arry_Header)
        
    'Declare Strings
    
    Dim strErrFullSource As String
    
    If str_ModuleName <> "" And str_SubName <> "" Then
        strErrFullSource = err_Source & "." & str_ModuleName & "." & str_SubName
    ElseIf str_ModuleName <> "" Then
        strErrFullSource = err_Source & "." & str_ModuleName
    ElseIf str_SubName <> "" Then
        strErrFullSource = err_Source & "." & str_SubName
    Else
        strErrFullSource = err_Source
    End If

' ------------------------------------
' Capture the change in the Change Log
' ------------------------------------

    With ws_ErrorLog
        .Cells(int_CurRow, col_ErrorOccured) = Format(Now, "m/d/yyyy hh:mm:ss")
        .Cells(int_CurRow, col_Errorer) = myFunctions.fx_Name_Reverse()
        .Cells(int_CurRow, col_ErrorSource) = strErrFullSource
        .Cells(int_CurRow, col_ErrorNum) = err_Num
        .Cells(int_CurRow, col_ErrorDesc) = err_Desc
        .Cells(int_CurRow, col_ProjectPath) = ThisWorkbook.FullName
    End With

' -------------------------
' Update the row formatting
' -------------------------

    Call fx_Steal_First_Row_Formating( _
        ws:=ws_ErrorLog, _
        intFirstRow:=2, _
        int_LastRow:=int_CurRow, _
        int_LastCol:=int_LastCol)

On Error GoTo 0

End Function
Function fx_Exception_User()

' Purpose: To output if the user is on the exception list or not.
' Trigger: Called
' Updated: 9/23/2020

' Change Log:
'          9/23/2020: Intial Creation
    
' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strUserID As String
        strUserID = Application.UserName

    Dim bolExceptionUser As Boolean

' -------------------------------------
' Determine if they user is on the list
' -------------------------------------

    If _
        strUserID = "Rinaldi, James" Or _
        strUserID = "Rauckhorst, Eric W." Or _
        strUserID = "Hogan, Elizabeth" Or _
        strUserID = "Barcikowski, Melissa H." Or _
        strUserID = "Renzulli, Scott W." _
    Then
        bolExceptionUser = True
    Else
        bolExceptionUser = False
    End If

    fx_Exception_User = bolExceptionUser

End Function
Function fx_Open_Workbook(strPromptTitle As String, Optional bol_ReadOnly As Boolean = True, Optional bol_CloseIfOpen As Boolean, Optional strFolderPath As String) As Workbook
             
' Purpose: This function will prompt the user for the workbook to open and returns that workbook.
' Trigger: Called Function
' Updated: 11/18/2023

' Change Log:
'       2/12/2021:  Initial Creation
'       2/12/2021:  Added the code to abort if the user selects cancel.
'       2/12/2021:  Added the code to determine if the Workbook is already open.
'       6/16/2021:  Added the code to ChDrive and ChDir
'       11/19/2022: Added the bol_CloseIfOpen value
'       11/17/2023: Added the strFolderPath optional variable and related code
'       11/18/2023: Updated to include the bol_ReadOnly Optional Variable (default of True)

' --------------------------------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Set wbTEST = fx_Open_Workbook(strPromptTitle:="Select the current Sageworks data dump")
    
' Use Example 2: _

' Legend:
'   bol_ReadOnly: If True then opens the workbook Read Only
'   bol_CloseIfOpen: If the workbook is already open then close it, and re-open
'   strFolderPath: If this is passed then include that part of the path

' ***********************************************************************************************************************************
             
On Error Resume Next
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path & strFolderPath
On Error GoTo 0
             
' -----------------
' Declare Variables
' -----------------
             
    Dim str_wbPath As String
        str_wbPath = Application.GetOpenFilename( _
        Title:=strPromptTitle, FileFilter:="Excel Workbooks (*.xls*;*.csv),*.xls*;*.csv")
             
        If str_wbPath = "False" Then
            MsgBox "No Workbook was selected, the code cannont continue."
            myPrivateMacros.DisableForEfficiencyOff
            End
        End If
        
' -----------------------------------------
' Determine if the Workbook is already open
' -----------------------------------------
        
    Dim bolAlreadyOpen As Boolean
        
     Dim str_WbName As String
         str_WbName = Right(str_wbPath, Len(str_wbPath) - InStrRev(str_wbPath, "\"))
        
    On Error Resume Next
        Dim wb As Workbook
        Set wb = Workbooks(str_WbName)
        bolAlreadyOpen = Not wb Is Nothing
    On Error GoTo 0
        
' -------------------------
' Set the Workbook variable
' -------------------------
        
    If bolAlreadyOpen = True Then
        If bol_CloseIfOpen = True Then
            Workbooks(str_WbName).Close savechanges:=False ' Added on 11/19/22
            Set fx_Open_Workbook = Workbooks.Open(str_wbPath, UpdateLinks:=False, ReadOnly:=bol_ReadOnly)
        Else
            Set fx_Open_Workbook = Workbooks(str_WbName)
        End If
        
    Else
        Set fx_Open_Workbook = Workbooks.Open(str_wbPath, UpdateLinks:=False, ReadOnly:=bol_ReadOnly)
    End If

End Function
Function fx_Select_Worksheet(str_WbName As String, str_WsName As String, Optional str_Ws_NamedRange As String) As Worksheet

' Purpose: To determine if a sheet exists, otherwise prompt the user for the sheet to select.
' Trigger: Called
' Updated: 6/5/2023
'
' Change Log:
'       7/4/2021:   Initial Creation, combined fx_Sheets_Exists and fx_Pick_Worksheet
'       6/5/2023:   Added the str_Ws_NamedRange optional varible to allow the user to update the ws Name
'
' -----------------------------------------------------------------------------------------------------------------------------------
'
' Use Example:
'    Dim wsSageworksRiskTrend As Worksheet
'    Set wsSageworksRiskTrend = fx_Select_Worksheet(wbSageworksChainedReport.Name, "1022 - Risk Trend 2873 ME Extra111")
'
'    Dim ws5003_Source As Worksheet
'    Set ws5003_Source = fx_Select_Worksheet( _
'        str_WbName:=wbSageworksChainedReport.Name, _
'        str_WsName:=ThisWorkbook.Worksheets(1).Evaluate("wsName_5003"), _
'        str_Ws_NamedRange:="wsName_5003")
'
' LEGEND OPTIONAL:
'   str_Ws_NamedRange: The name of the Named Range that corresponds to the applicable name of the worksheet.  If this is passed AND no sheet is found, update with the value selected by the user.
'
'
' ***********************************************************************************************************************************

' ---------------------------------
' Determine if the worksheet exists
' ---------------------------------

On Error GoTo SheetNotFound

    If Evaluate("ISREF('[" & str_WbName & "]" & str_WsName & "'!A1)") = True Then
        Set fx_Select_Worksheet = Workbooks(str_WbName).Sheets(str_WsName)
        Exit Function
    End If

SheetNotFound:

' -----------------------------------
' Prompt the user to select the sheet
' -----------------------------------

    MsgBox "The " & str_WsName & " was not found, please select the correct Worksheet."
        uf_wsSelector.Show
    
' ---------------------------
' Pass the selected worksheet
' ---------------------------
            
    If str_wsSelected = "" Then
        MsgBox "No Worksheet was selected, the code cannont continue."
        myPrivateMacros.DisableForEfficiencyOff
        End
    Else
        Set fx_Select_Worksheet = Workbooks(str_WbName).Worksheets(str_wsSelected)
        If str_Ws_NamedRange <> "" Then ThisWorkbook.Names(str_Ws_NamedRange).Value = str_wsSelected 'Updated the Named Range if it exists
    End If

End Function
Function fx_Sheet_Exists(str_WbName As String, str_WsName As String, Optional bol_DeleteSheet As Boolean) As Boolean

' Purpose: To determine if a sheet exists, to be used in an IF statement.
' Trigger: Called
' Updated: 7/26/2022

' Change Log:
'       6/29/2021:  Intial Creation
'       6/29/2021:  Added the ErrorHandler for the 2015 #VALUE error when the ws doesn't exist
'       7/26/2022:  Added the optional 'bol_DeleteSheet' and application.displayalert = false

' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
'    Dim bolSupportOpenAlready As Boolean
'    bolSupportOpenAlready = fx_Sheet_Exists( _
        str_WbName:=strProjName & ".xlsx", _
        str_WsName:="Next Actions")
'   If fx_Sheet_Exists(ThisWorkbook.Name, "VALIDATION") = False Then

' LEGEND:
'   bol_DeleteSheet: If the sheet exists delete it and then pass False

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler

Application.DisplayAlerts = False
    If Evaluate("ISREF('[" & str_WbName & "]" & str_WsName & "'!A1)") = True Then
        If bol_DeleteSheet = True Then
            Workbooks(str_WbName).Sheets(str_WsName).Delete
        Else
            fx_Sheet_Exists = True
        End If
    End If
Application.DisplayAlerts = True
        
    Exit Function

ErrorHandler:

fx_Sheet_Exists = False

End Function
Function fx_Create_Headers(str_Target_FieldName As String, arry_Target_Header As Variant) As Long

' Purpose: To determine the column number for a specific field in the header.
' Trigger: Called
' Updated: 7/3/2023

' Change Log:
'       5/1/2020: Intial Creation
'       12/11/2020: Updated to use an array instead of the range, reducing the time to run by 75%.
'       7/3/2023:   Switched to using a Dictionary which is ~50% faster, and easier to troubleshoot

' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    intColNum_Source = fx_Create_Headers_v2(str_Target_FieldName, arry_Target_Header)

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim i As Long

    Dim dict_LookupData As Scripting.Dictionary
    Set dict_LookupData = New Scripting.Dictionary
        dict_LookupData.CompareMode = TextCompare
        
    Dim arryTemp() As Variant
        arryTemp = arry_Target_Header

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------

On Error Resume Next
    For i = LBound(arryTemp) To UBound(arryTemp)
        dict_LookupData.Add Key:=arryTemp(i), Item:=i
    Next i
On Error GoTo 0

' --------------------------------------------------
' Loop through the array to find the matching column
' --------------------------------------------------

    If dict_LookupData.Exists(str_Target_FieldName) Then
        fx_Create_Headers = dict_LookupData.Item(str_Target_FieldName)
    End If

End Function
Function fx_Create_HeaderRow(varHeaderValues() As Variant, wsTarget As Worksheet, Optional intTargetHeaderRow As Long)
             
' Purpose: This function will create a row of headers based on the passed array of values in the wsTarget.
'
' Trigger: Called Function
' Updated: 11/18/2023
'
' Change Log:
'       11/18/2023: Initial Creation
'
' --------------------------------------------------------------------------------------------------------------------------------------------------------

' Use Example:
'    fx_Create_HeaderRow(varHeaderValues:="Select the current Sageworks data dump" wsTarget)
    
' Legend:
'   intHeaderRow: If passed use that instead of 1 as the HeaderRow

' ***********************************************************************************************************************************
        
' -----------------
' Declare Variables
' -----------------
        
    Dim intHeaderRow As Long
        intHeaderRow = WorksheetFunction.Max(1, intTargetHeaderRow)
        
    Dim rngHeader As Range
    Set rngHeader = wsTarget.Range(wsTarget.Cells(intHeaderRow, 1), wsTarget.Cells(intHeaderRow, UBound(varHeaderValues) + 1))
        
    Dim clrGrey As Long
        clrGrey = RGB(230, 230, 230)
        
' --------------------------------
' Add the values to the header row
' --------------------------------
            
    rngHeader.Value2 = varHeaderValues

' ------------------------------------------
' Customize the formatting of the header row
' ------------------------------------------

With rngHeader
    .Font.Bold = True
    .Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
    .Interior.Color = clrGrey
End With

    wsTarget.Range(intHeaderRow & ":" & intHeaderRow).AutoFilter

    wsTarget.Cells.EntireColumn.AutoFit

End Function
Function fx_Find_Row(ws As Worksheet, str_Target As String, Optional str_TargetFieldName As String, Optional str_TargetCol As String) As Long

' Purpose: To find the target value in the passed column for the passed worksheet.  Replaces the Find function, to account for hidden rows.
' Trigger: Called
' Updated: 3/6/2022

' Change Log:
'       12/26/2021: Initial Creation
'       12/27/2021: Made the int_LastRow more dynamic, and added the 1 to capture a blank row
'       1/19/2022:  Added the code to allow str_TargetCol to be passed
'       3/6/2022:   Added Error Handling around the Dictionary to allow duplicates
'                   Updated to handle situations where str_TargetCol AND str_TargetFieldName are not passed

' Note: Formerly called fx_Find_CurRow

' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Find_Row( _
        ws:=ThisWorkbook.Sheets("Projects"), _
        str_TargetFieldName:="Project", _
        str_Target:="P.343 - Migrate to Win10")

' Use Example 2: Passing the Target Field Name _
    intRowCurProject = fx_Find_Row(ws:=wsProjects, str_Target:=strProjName, str_TargetFieldName:="Project")

' Use Example 3: Passing the Target Column letter reference _
    intRowCurProject = fx_Find_Row(ws:=wsProjects, str_Target:=strProjName, str_TargetCol:="B")

' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    'Declare Header Variables
    
    Dim arry_Header_Data() As Variant
        arry_Header_Data = Application.Transpose(ws.Range(ws.Cells(1, 1), ws.Cells(1, 99)))
        
    Dim col_Target As Long
        If str_TargetCol <> "" Then
            col_Target = ws.Range(str_TargetCol & "1").Column
        ElseIf str_TargetFieldName <> "" Then
            col_Target = fx_Create_Headers(str_TargetFieldName, arry_Header_Data)
        Else
            col_Target = 1
        End If
    
    'Declare Other Variables

    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, col_Target).End(xlUp).Row, _
        ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row) + 1
        
    Dim arryData() As Variant
        arryData = ws.Range(ws.Cells(1, col_Target), ws.Cells(int_LastRow, col_Target))

    Dim dictData As New Scripting.Dictionary
        dictData.CompareMode = TextCompare
        
    Dim i As Long
        
' -------------------
' Fill the Dictionary
' -------------------
    
On Error Resume Next
    
    For i = 1 To UBound(arryData)
        dictData.Add Key:=arryData(i, 1), Item:=i
    Next i
    
On Error GoTo 0
    
' --------------------
' Find the Current Row
' --------------------
    
    fx_Find_Row = dictData(str_Target)

End Function
Function fx_Find_LastRow(ws_Target As Worksheet, _
Optional int_TargetColumn As Long, Optional bol_MinValue2 As Boolean, Optional bolIncludeSpecialCells As Boolean, Optional bolIncludeUsedRange As Boolean) As Long

' Purpose: To output the the Last Row for the passed ws using multiple options.
' Trigger: Called
' Updated: 11/19/2023
'
' Change Log:
'       11/29/2021: Initial Creation
'       3/6/2022:   Overhauled to include error handling, and the if statements to breakout the determination of the Last Row
'                   Added the fx_Find_Row code as an alternative to handle filtered data
'       11/18/2023: Fixed an unclosed error handler (missing On Error GoTo 0)
'                   Updated to simplify the code and remove an If statement
'       11/19/2023: Added the code for bol_MinValue2 to allow the option to not pass a # less than 2
'
' -----------------------------------------------------------------------------------------------------------------------------------
'
' Use Example: _
'   int_LastRow = fx_Find_LastRow(wsData)
'
' Use Example 2: Using all of the optional variables _
'   int_LastRow = fx_Find_LastRow(ws_Target:=wsTest, int_TargetColumn:=2, bolIncludeSpecialCells:=True, bolIncludeUsedRange:=True)
'
' Legend:
'   bolIncludeUsedRange: If this is True then the last row of the UsedRange will be included in the Max formula
'   bolIncludeSpecialCells: If this is True then the SpecialCells(xlCellTypeLastCell) row will be included in the Max formula
'
' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

On Error Resume Next
    
    ' Declare the 1st Option based on End(xlUp)
    
    Dim int_LastRow_1st As Long
    
    If int_TargetColumn <> 0 Then
        int_LastRow_1st = ws_Target.Cells(ws_Target.Rows.Count, int_TargetColumn).End(xlUp).Row
    Else
        int_LastRow_1st = ws_Target.Cells(ws_Target.Rows.Count, "A").End(xlUp).Row
    End If
    
    ' Declare the 2nd Option based on xlCellTypeLastCell
    
    If bolIncludeSpecialCells = True Then
        Dim int_LastRow_2nd As Long
            int_LastRow_2nd = ws_Target.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row
    End If
    
    ' Declare the 3rd Option based on UsedRange.Rows.Count
    
    If bolIncludeUsedRange = True Then
        Dim int_LastRow_3rd As Long
            int_LastRow_3rd = ws_Target.UsedRange.Rows(ws_Target.UsedRange.Rows.Count).Row
    End If
    
    ' Declare the 4th Option of a minimum of 2
    
    If bol_MinValue2 = True Then
        Dim int_LastRow_4th As Long
            int_LastRow_4th = 2
    End If
    
    ' Declare the Max Integer
    
    Dim int_LastRow_Max As Long

On Error GoTo 0

' ---------------------------------
' Determine which int_LastRow to use
' ---------------------------------

    int_LastRow_Max = WorksheetFunction.Max(int_LastRow_1st, int_LastRow_2nd, int_LastRow_3rd, int_LastRow_4th)
        
    fx_Find_LastRow = int_LastRow_Max
        
End Function

Function fx_Find_LastColumn(ws_Target As Worksheet, Optional bolIncludeSpecialCells As Boolean, Optional bolIncludeUsedRange As Boolean) As Long

' Purpose: To output the the Last Column for the passed ws using multiple options.
' Trigger: Called
' Updated: 11/18/2023
'
' Change Log:
'       3/6/2022:   Initial Creation, based on fx_Find_LastCol
'       11/18/2023: Fixed an unclosed error handler (missing On Error GoTo 0)
'                   Updated to simplify the code and remove an If statement
'
' -----------------------------------------------------------------------------------------------------------------------------------
'
' Use Example: _
'   int_LastCol = fx_Find_LastColumn(wsData)
'
' Legend:
'
'   bolIncludeUsedRange: If this is True then the last Col of the UsedRange will be included in the Max formula
'   bolIncludeSpecialCells: If this is True then the SpecialCells(xlCellTypeLastCell) col will be included in the Max formula
'
' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

On Error Resume Next

    ' Declare the 1st Option based on End(xlToLeft)
    
    Dim int_LastCol_1st As Long
        int_LastCol_1st = ws_Target.Cells(1, ws_Target.Columns.Count).End(xlToLeft).Column
        
    ' Declare the 2nd Option based on xlCellTypeLastCell
    
    If bolIncludeSpecialCells = True Then
        Dim int_LastCol_2nd As Long
            int_LastCol_2nd = ws_Target.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Column
    End If
     
    ' Declare the 3rd Option based on UsedRange.Rows.Count
    
    If bolIncludeUsedRange = True Then
        Dim int_LastCol_3rd As Long
            int_LastCol_3rd = ws_Target.UsedRange.Columns(ws_Target.UsedRange.Columns.Count).Column
    End If

    ' Declare the Max Integer
     
    Dim int_LastCol_Max As Long

On Error GoTo 0

' ---------------------------------
' Determine which int_LastCol to use
' ---------------------------------

    int_LastCol_Max = WorksheetFunction.Max(int_LastCol_1st, int_LastCol_2nd, int_LastCol_3rd, 2) 'Don't pass values <2
        
    fx_Find_LastColumn = int_LastCol_Max
        
End Function

Function fx_Copy_in_Data_for_Matching_Fields(wsSource As Worksheet, wsDest As Worksheet, _
 Optional int_Source_HeaderRow As Long, Optional int_LastRowtoImport As Long, Optional bol_AggregateData As Boolean, Optional bol_Steal1stRowFormatting As Boolean, _
 Optional bol_CloseSourceWb As Boolean, Optional bol_ImportVisibleFieldsOnly As Boolean, _
 Optional str_ModuleName As String, Optional str_ControlTotalField As String, Optional int_CurRow_wsValidation As Long)
    
' Purpose: To copy the data from the source to the destination, wherever the fields match.

' Trigger: Called
' Updated: 11/18/2023
'
' Change Log:
'       9/18/2020:  Intial Creation based on CV Mod Agg. CV Tracker import code
'       11/3/2020:  Updated to include the strSourceDesc and strDestDesc to feed into the validation function
'       11/3/2020:  Updated to include the str_ModuleName to feed into the validation function
'       11/3/2020:  Removed the 'DisableforEfficiency' as it was disabiling it in my Main Procedure.
'       2/12/2021:  Updated to account for pulling in only visible data
'       2/12/2021:  Switched from the filtered boolean to int_LastRowtoImport
'       2/12/2021:  Updated to use Arrays instead of Ranges for the import
'       2/12/2021:  Added the code related to bol_CloseSourceWb
'       2/25/2021:  Updated the code for str_ControlTotalField and str_ModuleName to make it optional
'       2/25/2021:  Removed the old col_Bal_Dest and col_Bal_Source references
'       3/9/2021:   Added the code to use intLastUsedRow_Dest to delete any extraneous rows
'       3/15/2021:  Updated to include the Optional intHeaderRow field to handle ignoring headers
'       5/17/2021:  Updated the code related to the Data Validation to use the intFirstRowData_Source instead of defaulting to 1
'       5/17/2021:  Updated the intHeaderRow variable to use the code from intHeaderRow_Source
'       6/16/2021:  Made some minor improvements to the variables to make them more resilient.
'       6/16/2021:  Added the code to apply the formatting from the first row to the rest.
'       6/21/2021:  Added Option to only import the visible fields
'       6/22/2021:  Updated the code to assign intHeaderRow_Source, and related code in the data
'       3/1/2022:   Added the int_CurRow_wsValidation variable to pass to fx_Create_Data_Validation_Control_Totals
'       5/5/2023:   Added a check for intLastUsedRow_Dest so that if it is 1 it changes to 2
'       5/23/2023:  Attempted to use an array but it failed when pulling in the L-SNB Acct #s (LIQ accounts starting in "=")
'                   Switched to using Copy => Paste for errors in the import
'       11/18/2023: Added the bol_AggregateData to allow me to aggregate the data, not remove the old
'                   Updated the int_LastRow to use fx_Find_LastRow and int_LastCol to use fx_Find_LastCol
'                   Created the int_lastDataRow_Dest variable if I am aggregating data
'                   Removed the 'If Err.Number = 7 Then 'Handle the LIQ Loan # errors' related code
'       11/19/2023: Created the bol_Steal1stRowFormatting to allow the formatting to be bypassed
'                   Simplified the loop related to bol_ImportVisibleFieldsOnly
'                   Update the Array code to result in one dimensional arrays

' -----------------------------------------------------------------------------------------------------------------------------------

'   EXAMPLE (SIMPLE): _
        Call fx_Copy_in_Data_for_Matching_Fields( _
            wsSource:=ws, _
            wsDest:=wsCombinedData, _
            bol_AggregateData:=True)

'   USE EXAMPLE: _
        Call fx_Copy_in_Data_for_Matching_Fields( _
            wsSource:=wsSource, _
            wsDest:=wsData, _
            int_Source_HeaderRow:=1, _
            int_LastRowtoImport:=0, _
            str_ModuleName:="o_11_Import_Data", _
            str_ControlTotalField:="New Direct Outstanding", _
            bol_CloseSourceWb:=True, _
            int_CurRow_wsValidation:=2, _
            bol_ImportVisibleFieldsOnly:=True)

' LEGEND MANDATORY:
'   wsSource:  The Source worksheet that the data is being copied FROM
'   wsDest:  The destination worksheet that the data is being copied TO

' LEGEND OPTIONAL:
'   int_Source_HeaderRow: The row that the header data is located, default is 1
'   int_LastRowtoImport: The last row of data to import, default is the max of LastRow for Col A, Col B, and Col C
'   bol_CloseSourceWb: Closes the parent workbook of the wsSource, so long as it isn't the workbook running the code
'   bol_ImportVisibleFieldsOnly: If set to True then imports only visibile fields in from the wsSource
'   str_ModuleName: Used to pass the module name of the module running the code to fx_Create_Data_Validation_Control_Totals
'   str_ControlTotalField: Used to pass the control total fiel name to fx_Create_Data_Validation_Control_Totals
'   int_CurRow_wsValidation: Used to pass the current row for the wsValidation to fx_Create_Data_Validation_Control_Totals
'   bol_AggregateData: If True then do NOT remove the old data in the destination

' ***********************************************************************************************************************************

' -----------------------------------------------
' Turn off any filtering from the source and dest
' -----------------------------------------------
        
    If wsSource.AutoFilterMode = True Then wsSource.AutoFilter.ShowAllData
        
    If wsDest.AutoFilterMode = True Then wsDest.AutoFilter.ShowAllData

' ------------------------
' Declare Source Variables
' ------------------------

    ' Declare Integers
    
    Dim int_LastRow_Source As Long
        If int_LastRowtoImport > 0 Then 'If I passed the int_LastRowtoImport variable use it
            int_LastRow_Source = int_LastRowtoImport
        Else
            int_LastRow_Source = fx_Find_LastRow(ws_Target:=wsSource, int_TargetColumn:=2, bolIncludeSpecialCells:=True, bolIncludeUsedRange:=True)
        End If

    Dim intHeaderRow_Source As Long
        
        If int_Source_HeaderRow > 0 Then 'If I passed the intHeaderRow variable use it
            intHeaderRow_Source = int_Source_HeaderRow
        Else
            intHeaderRow_Source = 1
        End If

    Dim intFirstRowData_Source As Long
        intFirstRowData_Source = intHeaderRow_Source + 1

    Dim int_LastCol_Source As Long
        int_LastCol_Source = fx_Find_LastColumn(wsSource)
        
' -----------------------------
' Declare Destination Variables
' -----------------------------

    ' Declare Integers
    
    Dim int_LastRow_Dest As Long
        int_LastRow_Dest = fx_Find_LastRow(ws_Target:=wsDest, int_TargetColumn:=2, bolIncludeSpecialCells:=True, bolIncludeUsedRange:=True)
    
    Dim int_CurRow_Dest As Long
        int_CurRow_Dest = int_LastRow_Dest + 1
        
    Dim intLastUsedRow_Dest As Long
        intLastUsedRow_Dest = wsDest.Range("A1").SpecialCells(xlCellTypeLastCell).Row
        
        If intLastUsedRow_Dest = 1 Then intLastUsedRow_Dest = 2
    
    Dim int_LastCol_Dest As Long
        int_LastCol_Dest = fx_Find_LastColumn(wsDest)
    
' -----------------------
' Declare Other Variables
' -----------------------
    
    ' Declare Other Integers
        
    Dim int_CurRowValidation As Long
        
    ' Declare "Ranges"
    
    Dim arry_Header_Source() As Variant
        arry_Header_Source = Application.Transpose(Application.Transpose( _
        wsSource.Range(wsSource.Cells(intHeaderRow_Source, 1), wsSource.Cells(intHeaderRow_Source, int_LastCol_Source))))
        
    Dim arry_Header_Dest() As Variant
        arry_Header_Dest = Application.Transpose(Application.Transpose( _
        wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, int_LastCol_Dest)).Value2))
        
    ' Declare Values for Loops
        
    Dim strFieldName As String
    
    Dim intColNum_Source As Long
    
    Dim intColNum_Dest As Long
        
    Dim i As Long
    
    Dim int_lastDataRow_Dest As Long
        int_lastDataRow_Dest = WorksheetFunction.Max _
        ((int_LastRow_Source - intHeaderRow_Source + int_CurRow_Dest - 1), (int_CurRow_Dest))
    
' ---------------------------------
' Declare Data Validation Variables
' ---------------------------------
    
    ' Declare Strings
    
    Dim strSourceDesc As String
        strSourceDesc = wsSource.Parent.Name & " - " & wsSource.Name

    Dim strDestDesc As String
        strDestDesc = wsDest.Parent.Name & " - " & wsDest.Name

' ------------------------------------
' Clear out the old data and cell fill
' ------------------------------------
    
    If bol_AggregateData = False Then
        wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(int_LastRow_Dest, int_LastCol_Dest)).ClearContents
        wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(int_LastRow_Dest, int_LastCol_Dest)).Interior.Color = xlNone
        wsDest.Range(wsDest.Cells(int_LastRow_Dest + 1, 1), wsDest.Cells(intLastUsedRow_Dest, 1)).EntireRow.Delete
    End If

' ----------------------------------
' Copy over the data from the source
' ----------------------------------

    For intColNum_Dest = 1 To int_LastCol_Dest
        
        strFieldName = arry_Header_Dest(intColNum_Dest)
        intColNum_Source = fx_Create_Headers(strFieldName, arry_Header_Source)
                
        If intColNum_Source > 0 Then
            If bol_ImportVisibleFieldsOnly = True And wsDest.Columns(intColNum_Dest).Hidden = False Then
                wsDest.Range(wsDest.Cells(int_CurRow_Dest, intColNum_Dest), wsDest.Cells(int_lastDataRow_Dest, intColNum_Dest)).Value2 = _
                wsSource.Range(wsSource.Cells(intFirstRowData_Source, intColNum_Source), wsSource.Cells(int_LastRow_Source, intColNum_Source)).Value2
            End If
            
            If bol_ImportVisibleFieldsOnly = False Then
                wsDest.Range(wsDest.Cells(int_CurRow_Dest, intColNum_Dest), wsDest.Cells(int_lastDataRow_Dest, intColNum_Dest)).Value2 = _
                wsSource.Range(wsSource.Cells(intFirstRowData_Source, intColNum_Source), wsSource.Cells(int_LastRow_Source, intColNum_Source)).Value2
            End If
        End If

    Next intColNum_Dest

    ' Reset the Last Row variable
    int_LastRow_Dest = fx_Find_LastRow(ws_Target:=wsDest, int_TargetColumn:=2, bolIncludeSpecialCells:=True, bolIncludeUsedRange:=True)

' ------------------------------------------------------------
' Output the control totals to the Validation ws, if it exists
' ------------------------------------------------------------
    
    If str_ControlTotalField <> "" Then
        
        Call fx_Create_Data_Validation_Control_Totals( _
            wsDataSource:=wsSource, _
            str_ModuleName:=str_ModuleName, _
            strSourceName:=strSourceDesc, _
            intHeaderRow:=intHeaderRow_Source, _
            int_LastRowtoImport:=int_LastRowtoImport, _
            str_ControlTotalField:=str_ControlTotalField, _
            int_CurRow_wsValidation:=int_CurRow_wsValidation)
        
        Call fx_Create_Data_Validation_Control_Totals( _
            wsDataSource:=wsDest, _
            str_ModuleName:=str_ModuleName, _
            strSourceName:=strDestDesc, _
            intHeaderRow:=1, _
            str_ControlTotalField:=str_ControlTotalField, _
            int_CurRow_wsValidation:=int_CurRow_wsValidation + 1)
        
    End If
    
' -----------------------------------------------
' Apply the formatting to all rows from the first
' -----------------------------------------------
    
    If bol_Steal1stRowFormatting = True Then
        
        Call fx_Steal_First_Row_Formating( _
            ws:=wsDest, _
            intFirstRow:=2, _
            int_LastRow:=int_LastRow_Dest, _
            int_LastCol:=int_LastCol_Dest)
        
    End If
    
' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function

Function fx_Create_Data_Validation_Control_Totals(wsDataSource As Worksheet, str_ModuleName As String, strSourceName As String, intHeaderRow As Long, str_ControlTotalField As String, Optional int_LastRowtoImport As Long, Optional int_CurRow_wsValidation As Long, Optional dblTotalsFromSource As Double, Optional intRecordCountFromSource As Long)

' Purpose: To output the data validation control totals to the wsValidation, if it exists.
' Trigger: Called
' Updated: 1/5/2022


' Change Log:
'       9/26/2020: Initial Creation
'       11/3/2020: Updated to activate ThisWorkbook before checking for the Validation ws
'       12/19/2020: Made the intRecordCount more resiliant
'       12/19/2020: Added the ThisWorkbook.Name to the ISREF check
'       2/12/2021: Added the code for int_LastRowtoImport
'       5/17/2021: Updated the calculation for strRng_Totals to go down to intRecordCount + intHeaderRow
'       6/16/2021: Added the code related to int_CurRow_wsValidation
'       6/22/2021: Updated the code for intRecordCount
'       6/22/2021: Updated the Check for the Validation worksheet to reference ThisWorkbook, and avoid the .Activate
'       1/5/2022:  Added the code to bypass the sums and whatnot if dblTotalsFromSource or intRecordCountFromSource is passed
    
' -----------------------------------------------------------------------------------------------------------------------------------
    
' Use Example: _
    Call fx_Create_Data_Validation_Control_Totals( _
        wsDataSource:=wsDest, _
        str_ModuleName:=str_ModuleName, _
        strSourceName:=strDestDesc, _
        intHeaderRow:=1, _
        str_ControlTotalField:=str_ControlTotalField, _
        int_CurRow_wsValidation:=3)

' Use Example 2: _
    Call fx_Create_Data_Validation_Control_Totals( _
    wsDataSource:=ws3666_Source, _
    str_ModuleName:="o3_Import_BB_Data", _
    strSourceName:=ws3666_Source.Parent.Name & " - " & ws3666_Source.Name, _
    intHeaderRow:=intHeaderRow, _
    str_ControlTotalField:="Line Commitment", _
    int_CurRow_wsValidation:=10, _
    intRecordCountFromSource:=intDestRowCounter - 2, _
    dblTotalsFromSource:=dblControlTotal)
    
' ***********************************************************************************************************************************

    'Only run of the VALIDATION ws exists
    If Evaluate("ISREF(" & "'[" & ThisWorkbook.Name & "]" & "Validation'" & "!A1)") = False Then
        Debug.Print "fx_Create_Data_Validation_Control_Totals failed becuase there is no ws called 'VALIDATION' in the Workbook"
        Exit Function
    End If

' ----------------------------
' Declare Validation Variables
' ----------------------------

    'Dim Worksheets

    Dim wsValidation As Worksheet
    Set wsValidation = ThisWorkbook.Sheets("VALIDATION")

    Dim wsSource As Worksheet
    Set wsSource = wsDataSource

    'Dim Cell References

    Dim int_LastCol As Long
        int_LastCol = wsSource.Cells(intHeaderRow, Columns.Count).End(xlToLeft).Column
      
    Dim int_CurRow As Long
        If int_CurRow_wsValidation > 1 Then 'If I passed the int_CurRow_wsValidation variable use it
            int_CurRow = int_CurRow_wsValidation
        Else
            int_CurRow = wsValidation.Cells(Rows.Count, "A").End(xlUp).Row + 1
        End If

    'Bypass the code if dblTotals was passed
    If dblTotalsFromSource > 0 Or intRecordCountFromSource > 0 Then GoTo Bypass

' ------------------------
' Declare Source Variables
' ------------------------

    'Dim "Ranges"
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(wsSource.Range(wsSource.Cells(intHeaderRow, 1), wsSource.Cells(intHeaderRow, int_LastCol)))
        
    Dim intColTotals As Long
        intColTotals = fx_Create_Headers(str_ControlTotalField, arry_Header)
    
    'Dim Integers

    Dim intRecordCount As Long
        If int_LastRowtoImport > 0 Then 'If I passed the int_LastRowtoImport variable use it
            intRecordCount = int_LastRowtoImport - intHeaderRow
        Else
            intRecordCount = WorksheetFunction.Max( _
            wsSource.Cells(Rows.Count, "A").End(xlUp).Row, _
            wsSource.Cells(Rows.Count, "B").End(xlUp).Row, _
            wsSource.Cells(Rows.Count, "C").End(xlUp).Row) - intHeaderRow
        End If
    
    'Dim Other Variables

    Dim strCol_Totals As String
        strCol_Totals = Split(Cells(1, intColTotals).Address, "$")(1)
    
    Dim strRng_Totals As String
        strRng_Totals = strCol_Totals & "1:" & strCol_Totals & intRecordCount + intHeaderRow
        
    Dim dblTotals As Double
        dblTotals = Round(Application.WorksheetFunction.Sum(wsSource.Range(strRng_Totals)), 2)

Bypass:

    'Assign the Optional variables if they were passed
    If dblTotalsFromSource > 0 Then dblTotals = dblTotalsFromSource
    If intRecordCountFromSource > 0 Then intRecordCount = intRecordCountFromSource

' ------------------------------------------------------
' Output the validation totals from the passed variables
' ------------------------------------------------------

    With wsValidation
        .Range("A" & int_CurRow) = Format(Now, "m/d/yyyy hh:mm")   'Date / Time
        .Range("B" & int_CurRow) = str_ModuleName                   'Code Module
        .Range("C" & int_CurRow) = strSourceName                   'Source
        .Range("D" & int_CurRow) = Format(dblTotals, "$#,##0")     'Total
        .Range("E" & int_CurRow) = Format(intRecordCount, "0,0")   'Count
    End With

End Function
Function fx_Validate_Control_Totals(int1stTotalRow As Long, int2ndTotalRow As Long)

' Purpose: To validate that the control totals for the data imported match.
' Trigger: Called
' Updated: 6/8/2022

' Use Example: _
    Call fx_Validate_Control_Totals(int1stTotalRow:=2, int2ndTotalRow:=3)

' Change Log:
'       6/22/2021:  Intial Creation
'       1/31/2022:  Updated to switch to an Information msgbox
'       6/8/2022:   Updated the control total formatting for the count not matching

' ***********************************************************************************************************************************

    ' Only run of the VALIDATION ws exists
    If Evaluate("ISREF(" & "'[" & ThisWorkbook.Name & "]" & "Validation'" & "!A1)") = False Then
        Debug.Print "fx_Validate_Control_Totals failed becuase there is no ws called 'VALIDATION' in the Workbook"
        Exit Function
    End If

' -----------
' Declare your variables
' -----------
    
    ' Dim Worksheets
    
    Dim wsValidation As Worksheet
    Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
    
    ' Dim Integers
    
    Dim int1stTotal As Double
        int1stTotal = wsValidation.Cells(int1stTotalRow, "D").Value
    
    Dim int2ndTotal As Double
        int2ndTotal = wsValidation.Cells(int2ndTotalRow, "D").Value
    
    Dim int1stCount As Long
        int1stCount = wsValidation.Cells(int1stTotalRow, "E").Value
    
    Dim int2ndCount As Long
        int2ndCount = wsValidation.Cells(int2ndTotalRow, "E").Value
        
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
        Buttons:=vbOKOnly + vbInformation, _
        Prompt:="The validation totals match, you're golden. " & Chr(10) & Chr(10) _
        & "1st Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "1st Validation Count: " & Format(int1stCount, "0,0")
       
    ElseIf bolTotalsMatch = False Then
    MsgBox Title:="Validation Totals Don't Match", _
        Buttons:=vbCritical, _
        Prompt:="The validation totals from the workbook don't match what was imported. " _
        & "Please review the totals in the Validation worksheet to determine what went awry. " & Chr(10) & Chr(10) _
        & "1st Validation Total Variance: " & Format(int1stTotal - int2ndTotal, "$#,##0") & Chr(10) & Chr(10) _
        & "1st Validation Count Variance: " & Format(int1stCount - int2ndCount, "0,0")

    End If
    
End Function
Function fx_Record_Change_in_DataChangeLog(strFieldChanged As String, strPriorValue As String, strNewValue As String, Optional strBorrowerLOB As String, Optional strBorrowerName As String, Optional strChangeType As String, Optional strChangeSource As String)

' Purpose: To record a change in the Change Log.
' Trigger: Workbook Change
' Updated: 2/19/2022

' Change Log:
'       12/1/2021:  Intial Creation
'       12/3/2021:  Added the 'fx_Steal_First_Row_Formating' code
'       12/6/2021:  Added the code to capture seconds
'       12/24/2021: Updated the name of the Change Log -> "DATA CHANGE LOG"
'       1/19/2022:  Added the IF statements for if the ChangeType and ChangeSource don't exist
'       2/19/2022:  Updated to use int_CurRow in the fx_Steal_Formatting
'                   Made the strBorrowerName as Optional

' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Record_Change_in_DataChangeLog( _
        strBorrowerName:="Rose Apothecary", _
        strBorrowerLOB:="TBD", _
        strFieldChanged:="RCCRP", _
        strPriorValue:=6, _
        strNewValue:=7, _
        strChangeType:="Credit Risk Change (Orange)", _
        strChangeSource:="PAR Meeting")

' Use Example 2: _
    Call fx_Record_Change_in_DataChangeLog( _
        strFieldChanged:=rngTargetCell.Worksheet.Cells(1, rngTargetCell.Column), _
        strPriorValue:=strOldValue, _
        strNewValue:=strNewValue)

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare Worksheets

    Dim wsDataChangeLog As Worksheet
    Set wsDataChangeLog = ThisWorkbook.Sheets("DATA CHANGE LOG")

    'Declare Integers

    Dim int_CurRow As Long
        int_CurRow = wsDataChangeLog.Range("A:A").Find("").Row
        
    Dim int_LastCol As Long
        int_LastCol = WorksheetFunction.Max( _
           wsDataChangeLog.Cells(1, Columns.Count).End(xlToLeft).Column, _
           wsDataChangeLog.Rows(1).Find("").Column - 1)
        
    'Declare Cell References
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(wsDataChangeLog.Range(wsDataChangeLog.Cells(1, 1), wsDataChangeLog.Cells(1, int_LastCol)))

    Dim col_ChangeMade As Long
        col_ChangeMade = fx_Create_Headers("Change Made", arry_Header)

    Dim col_Changer As Long
        col_Changer = fx_Create_Headers("By Who", arry_Header)

    Dim col_LOB As Long
        col_LOB = fx_Create_Headers("Line of Business", arry_Header)

    Dim col_Borrower As Long
        col_Borrower = fx_Create_Headers("Borrower", arry_Header)

    Dim col_FieldChanged As Long
        col_FieldChanged = fx_Create_Headers("Field Changed", arry_Header)

    Dim col_PriorValue As Long
        col_PriorValue = fx_Create_Headers("Prior Value", arry_Header)

    Dim col_NewValue As Long
        col_NewValue = fx_Create_Headers("New Value", arry_Header)

    Dim col_ChangeType As Long
        col_ChangeType = fx_Create_Headers("Change Type", arry_Header)

    Dim col_ChangeSource As Long
        col_ChangeSource = fx_Create_Headers("Change Source", arry_Header)

' ------------------------------------
' Capture the change in the Change Log
' ------------------------------------

    With wsDataChangeLog
        .Cells(int_CurRow, col_ChangeMade) = Format(Now, "m/d/yyyy hh:mm:ss")
        .Cells(int_CurRow, col_Changer) = myFunctions_Specific.fx_Name_Reverse()
        .Cells(int_CurRow, col_LOB) = strBorrowerLOB
        .Cells(int_CurRow, col_Borrower) = strBorrowerName
        .Cells(int_CurRow, col_FieldChanged) = strFieldChanged
        .Cells(int_CurRow, col_PriorValue) = strPriorValue
        .Cells(int_CurRow, col_NewValue) = strNewValue
        If strChangeType <> "" Then .Cells(int_CurRow, col_ChangeType) = strChangeType
        If strChangeSource <> "" Then .Cells(int_CurRow, col_ChangeSource) = strChangeSource
    End With

' -------------------------
' Update the row formatting
' -------------------------

    Call fx_Steal_First_Row_Formating( _
        ws:=wsDataChangeLog, _
        intSingleRow:=curRow)

End Function
Function fx_Capture_Change_for_DataChangeLog(rngTargetCell As Range, Optional strLOBFieldName As String, Optional strBorrowerFieldName As String)

' Purpose: To capture a change to the passed Target cell for the Change Log.
' Trigger: Worksheet Change Event
' Updated: 3/2/2022

' Change Log:
'       2/19/2022:  Initial Creation, based on CV Tracker code
'       3/2/2022:   Updated to pass the LOB and Borrower names to the fx_Record_Change_in_DataChangeLog

' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Capture_Change_for_DataChangeLog(rngTargetCell:=Target, strLOBFieldName:="Segment (LOB)", strBorrowerFieldName:="Borrower Name")

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler
            
' -----------------
' Declare Variables
' -----------------
    
    'Declare Strings
    Dim strOldValue As String
    
    Dim strNewValue As String
    
    'Declare Worksheets
    Dim ws_Target As Worksheet
    Set ws_Target = rngTargetCell.Parent
    
    'Declare Field References
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(ws_Target.Range(ws_Target.Cells(1, 1), ws_Target.Cells(1, ws_Target.Cells(1, Columns.Count).End(xlToLeft).Column)))
    
    Dim col_LOB As Long
        col_LOB = fx_Create_Headers(strLOBFieldName, arry_Header)

    Dim col_Borrower As Long
        col_Borrower = fx_Create_Headers(strBorrowerFieldName, arry_Header)
    
' ------------------------------
' Capture the old and new values
' ------------------------------

Application.ScreenUpdating = False

    strNewValue = rngTargetCell.Value

    Application.EnableEvents = False
        Application.Undo
        strOldValue = rngTargetCell.Value
        Application.Undo
    Application.EnableEvents = True

' --------------------------------------
' Copy the data into the Data Change Log
' --------------------------------------

    If strBorrowerFieldName <> "" Then
    
        Call fx_Record_Change_in_DataChangeLog( _
            strBorrowerLOB:=ws_Target.Cells(rngTargetCell.Row, col_LOB), _
            strBorrowerName:=ws_Target.Cells(rngTargetCell.Row, col_Borrower), _
            strFieldChanged:=rngTargetCell.Worksheet.Cells(1, rngTargetCell.Column), _
            strPriorValue:=strOldValue, _
            strNewValue:=strNewValue)

    Else

        Call fx_Record_Change_in_DataChangeLog( _
            strFieldChanged:=rngTargetCell.Worksheet.Cells(1, rngTargetCell.Column), _
            strPriorValue:=strOldValue, _
            strNewValue:=strNewValue)
            
    End If

Application.ScreenUpdating = True

Exit Function

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Function

Function fx_Update_Single_Field(wsSource As Worksheet, wsDest As Worksheet, _
    str_Source_TargetField As String, str_Source_MatchField As String, str_Dest_TargetField As String, str_Dest_MatchField As String, _
    Optional int_SourceHeaderRow As Long, Optional bol_ConvertMatchSourcetoValues As Boolean, _
    Optional bol_CloseSourceWb As Boolean, Optional bol_SkipDuplicates As Boolean, Optional bol_BlanksOnly As Boolean, _
    Optional str_OnlyUseValue As String, Optional arry_OnlyUseMultipleValues As Variant, Optional bol_MultipleOnlyUseValues As Boolean, _
    Optional bol_MissingLookupData_MsgBox As Boolean, Optional bol_MissingLookupData_UseExistingData As Boolean, _
    Optional strMissingLookupData_ValuetoUse, Optional str_WsNameLookup As String, _
    Optional str_FilterField_Dest As String, Optional str_FilterValue As String, Optional bol_FilterPassArray As Boolean)

' Purpose: To update the data in the Target Field in the Destination, based on data from the Target Field in the Source.
    ' 1) Add the data to be referenced into the dictionary
    ' 2) Look for the data to be updated for matches to the reference dictionary
    ' 3) Replace the existing data with the reference data

' Trigger: Called
' Updated: 6/29/2023

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
'       4/19/2022:  Added 'strMissingLookupData_ValuetoUse' to allow a user to pass a value that will be used for blanks
'                   Updated the names of some of the variables to help clarify
'       6/13/2022:  Added the 'str_WsNameLookup' and related code when a value is missing.
'                   Added code to remove the leading line break in str_missingvalues
'       9/15/2022:  Added the 'Or InStr(1, str_FilterValue, arry_Dest_Filter(i)) > 0' to allow an 'array' to be passed as a criteria
'                   Simplified how the Arrays are determined
'       6/13/2023:  Added a simple example and some clarifications
'       6/29/2023:  Updated to allow multiple strings to be passed to str_OnlyUseValue and arryTemp_OnlyUseValue to determine if the value is in the data
'                   Cleaned up the code for multiple strings to be passed

' -----------------------------------------------------------------------------------------------------------------------------------

' USE EXAMPLE 1 (BASIC): _
    Call fx_Update_Single_Field( _
        wsSource:=wsLists_Collateral, wsDest:=wsData, _
        str_Source_TargetField:="Collateral Bucket", _
        str_Source_MatchField:="Webster Collateral Type", _
        str_Dest_TargetField:="Collateral Bucket", _
        str_Dest_MatchField:="Collateral Type")

' USE EXAMPLE 2 (Missing Data Warning): _
    Call fx_Update_Single_Field( _
        wsSource:=wsLists_Collateral, wsDest:=wsData, _
        str_Source_TargetField:="Collateral Bucket", _
        str_Source_MatchField:="Webster Collateral Type", _
        str_Dest_TargetField:="Collateral Bucket", _
        str_Dest_MatchField:="Collateral Type", _
        bol_MissingLookupData_MsgBox:=True, _
        strWsNameLookup:="Collateral Lookup - Webster Collateral Type / Collateral Bucket", _
        strMissingLookupData_ValuetoUse:="Unidentified")

' USE EXAMPLE 3: _
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

' USE EXAMPLE 4 (Pass Array): _
    Call fx_Update_Single_Field( _
        wsSource:=wsPolicyExceptions_LL, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="1 - Collateral (#)", _
        str_Source_MatchField:="Helper", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Helper", _
        str_FilterField_Dest:="Line of Business", _
        str_FilterValue:="Asset Based Lending, Commercial Real Estate, Middle Market Banking, Sponsor and Specialty Finance, Wealth", _
        bol_FilterPassArray:=True)

' LEGEND MANDATORY:
'   wsSource:  The Source worksheet that the data is being copied FROM
'   wsDest:  The destination worksheet that the data is being copied TO
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

'   str_OnlyUseValue: Used to allow a SINGLE value to be used
'   arry_OnlyUseMultipleValues: Used to allow MULTIPLE values to be used

'   bol_MissingLookupData_MsgBox: Outputs a message box with a list of fields that are missing from the lookup
'   bol_MissingLookupData_UseExistingData: Will use the existing data instead of the lookup value
'   strMissingLookupData_ValuetoUse: If the value isn't in the lookup, and I didn't include a blank in the lookups, will use this value instead
'   str_WsNameLookup: If the value isn't in the lookup it will say where it was looking to help with troubleshooting
'   str_FilterField_Dest: Used to filter down the values to be imported on
'   str_FilterValue: Used to filter down the values to be imported on
'   bol_FilterPassArray: Allows an array of values to be passed instead of a single value for the filter

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
        
    ' Dim Arrays
    
    Dim arry_Dest_Target() As Variant
        arry_Dest_Target = Application.Transpose(.Range(.Cells(1, col_Dest_Target), .Cells(int_LastRow_wsDest, col_Dest_Target)))
    
    Dim arry_Dest_Match() As Variant
        arry_Dest_Match = Application.Transpose(.Range(.Cells(1, col_Dest_Match), .Cells(int_LastRow_wsDest, col_Dest_Match)))

    Dim arry_Dest_Filter() As Variant
        arry_Dest_Filter = Application.Transpose(.Range(.Cells(1, col_Dest_FilterField), .Cells(int_LastRow_wsDest, col_Dest_FilterField)))
 
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
    
    Dim str_OnlyUseMultipleValues As String
        If bol_MultipleOnlyUseValues = True Then
            str_OnlyUseMultipleValues = Join(arry_OnlyUseMultipleValues, ", ")
        End If
    
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

            ElseIf bol_MultipleOnlyUseValues = True Then
                If InStr(1, str_OnlyUseMultipleValues, arry_Source_Target(i)) > 0 Then ' Only import if the target matches the passed values
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
        
        If str_FilterField_Dest = "" Or arry_Dest_Filter(i) = str_FilterValue Or bol_FilterPassArray = True And InStr(1, str_FilterValue, arry_Dest_Filter(i)) > 0 Then
        
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
            
                ' If I have a record for a blank in the lookups use that, or use the strMissingLookupData_ValuetoUse if that was passed, otherwise abort
                If dict_LookupData.Exists(" ") = True Then
                    arry_Dest_Target(i) = dict_LookupData.Item(" ")
                
                ElseIf IsMissing(strMissingLookupData_ValuetoUse) = False Then
                    arry_Dest_Target(i) = strMissingLookupData_ValuetoUse
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
Function fx_Sum_or_Count_on_Single_Field(wsSource As Worksheet, wsDest As Worksheet, str_Source_TargetField As String, str_Source_MatchField As String, str_Dest_TargetField As String, str_Dest_MatchField As String, str_SumCriteria As String, opt_SumOrCount As opt_fx_Sum_or_Count_on_Single_Field, Optional strSumField As String, Optional bolNotBlankSum As Boolean, Optional bol_CloseSourceWb As Boolean)

' Purpose: To import the summed / counted data from the Source to the Destination.
    ' 1) Add the data to be referenced into the dictionary
    ' 2) Look for the data to be updated for matches to the reference dictionary
    ' 3) Replace the existing data with the reference data

' Trigger: Called
' Updated: 11/16/2021

' Change Log:
'       6/29/2021: Initial Creation, based on fx_Update_Single_Field
'       6/29/2021: Added the 'And arry_Target_Source(i - 1) = str_SumCriteria' to ensure _
                    we were comparing the same Exception AND the same Account Number
'       11/16/2021: Added the strSumField
'       9/15/2022:  Created the bolNotBlankSum boolean to allow me to sum when not blank

' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example 1: _
    Call fx_Sum_or_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="14 Digit Account Number", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Account Number", _
        str_SumCriteria:="1 - Collateral & Other Support", _
        opt_SumOrCount:=Count_Values)

' Use Example 2: _
    Call fx_Sum_or_Count_on_Single_Field( _
        wsSource:=ws5010a_LL, wsDest:=wsSageworksRT, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="14 Digit Account Number", _
        str_Dest_TargetField:="1 - Collateral & Other Support", _
        str_Dest_MatchField:="Account Number", _
        str_SumCriteria:="1 - Collateral & Other Support", _
        strSumField:="Risk Exposure (Loan Level)", _
        opt_SumOrCount:=Sum_Values)

' LEGEND MANDATORY:
'   TBD:

' LEGEND OPTIONAL:
'   bolNotBlankSum: When the value in the terget isn't blank then do the Sum

' ***********************************************************************************************************************************

' -------------------------------
' Declare your wsSource variables
' -------------------------------
        
With wsSource
        
    ' Dim Integers
    
    Dim int_LastCol_wsSource As Long
        int_LastCol_wsSource = .Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim int_LastRow_wsSource As Long
        int_LastRow_wsSource = WorksheetFunction.Max( _
        .Cells(Rows.Count, "A").End(xlUp).Row, _
        .Cells(Rows.Count, "B").End(xlUp).Row, _
        .Cells(Rows.Count, "C").End(xlUp).Row)
        
    ' Dim "Ranges"
        
    Dim arry_Header_wsSource() As Variant
        arry_Header_wsSource = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsSource)))
        
    Dim col_Target_Source As Integer
        col_Target_Source = fx_Create_Headers(str_Source_TargetField, arry_Header_wsSource)

    Dim col_Match_Source As Integer
        col_Match_Source = fx_Create_Headers(str_Source_MatchField, arry_Header_wsSource)
        
    Dim col_SumField_Source As Integer
        col_SumField_Source = fx_Create_Headers(strSumField, arry_Header_wsSource)
        If strSumField = "" Then col_SumField_Source = 1

    ' Dim Ranges
        
    Dim rng_Target_Source As Range
    Set rng_Target_Source = .Range(.Cells(1, col_Target_Source), .Cells(int_LastRow_wsSource, col_Target_Source))
        
    Dim rng_SumField_Source As Range
    Set rng_SumField_Source = .Range(.Cells(1, col_SumField_Source), .Cells(int_LastRow_wsSource, col_SumField_Source))
        
    Dim rng_Match_Source As Range
    Set rng_Match_Source = .Range(.Cells(1, col_Match_Source), .Cells(int_LastRow_wsSource, col_Match_Source))
        
    ' Dim Arrays
    
    Dim arry_Target_Source() As Variant
        arry_Target_Source = Application.Transpose(rng_Target_Source)
    
    Dim arry_SumField_Source() As Variant
        arry_SumField_Source = Application.Transpose(rng_SumField_Source)
    
    Dim arry_Match_Source() As Variant
        arry_Match_Source = Application.Transpose(rng_Match_Source)
        
End With
        
' -----------------------------
' Declare your wsDest variables
' -----------------------------
        
With wsDest
        
    ' Dim wsDest Integers
    
    Dim int_LastCol_wsDest As Long
        int_LastCol_wsDest = wsDest.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim int_LastRow_wsDest As Long
        int_LastRow_wsDest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "B").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "C").End(xlUp).Row)
 
    ' Dim wsDest "Ranges"
    
    Dim arry_Header_wsDest() As Variant
        arry_Header_wsDest = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsDest)))
        
    Dim col_Target_Dest As Integer
        col_Target_Dest = fx_Create_Headers(str_Dest_TargetField, arry_Header_wsDest)

    Dim col_Match_Dest As Integer
        col_Match_Dest = fx_Create_Headers(str_Dest_MatchField, arry_Header_wsDest)
 
    ' Dim Ranges
        
    Dim rng_Target_Dest As Range
    Set rng_Target_Dest = .Range(.Cells(1, col_Target_Dest), .Cells(int_LastRow_wsDest, col_Target_Dest))
        
    Dim rng_Match_Dest As Range
    Set rng_Match_Dest = .Range(.Cells(1, col_Match_Dest), .Cells(int_LastRow_wsDest, col_Match_Dest))
        
    ' Dim Arrays
    
    Dim arry_Target_Dest() As Variant
        arry_Target_Dest = Application.Transpose(rng_Target_Dest)
    
    Dim arry_Match_Dest() As Variant
        arry_Match_Dest = Application.Transpose(rng_Match_Dest)
 
End With
        
' ----------------------------
' Declare your Other variables
' ----------------------------
        
    ' Dim Dictionaries
    
    Dim dict_LookupData As Scripting.Dictionary
        Set dict_LookupData = New Scripting.Dictionary
        dict_LookupData.CompareMode = TextCompare
        
    ' Declare Loop Variables
    
    Dim i As Long
        
    Dim dblSum As Double
    
    Dim intCount As Long

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------
    
On Error Resume Next
       
        For i = 2 To UBound(arry_Target_Source)
           ' If i = 55 Then Stop
            If arry_Target_Source(i) = str_SumCriteria Or bolNotBlankSum = True And arry_Target_Source(i) <> "" Then
                If arry_Match_Source(i) = arry_Match_Source(i - 1) And arry_Target_Source(i - 1) = str_SumCriteria Then
                    If opt_SumOrCount = Sum_Values Then
                        dblSum = dblSum + arry_SumField_Source(i)
                    Else
                        intCount = intCount + 1
                    End If
                Else
                    If opt_SumOrCount = Sum_Values Then
                        dblSum = dblSum + arry_SumField_Source(i)
                        dict_LookupData.Add Key:=arry_Match_Source(i), Item:=dblSum
                        dblSum = 0
                    Else
                        intCount = intCount + 1
                        dict_LookupData.Add Key:=arry_Match_Source(i), Item:=intCount
                        intCount = 0
                    End If
                End If
            
                ' To handle the last record
                If i = UBound(arry_Target_Source) Then
                    If opt_SumOrCount = Sum_Values Then
                        dblSum = dblSum + arry_SumField_Source(i)
                        dict_LookupData.Add Key:=arry_Match_Source(i), Item:=dblSum
                        dblSum = 0
                    Else
                        intCount = intCount + 1
                        dict_LookupData.Add Key:=arry_Match_Source(i), Item:=intCount
                        intCount = 0
                    End If
        
                End If
            
            End If
        
        Next i
        
On Error GoTo 0
    
' -----------------------------------------------------------------------------
' Loop through the Lookup Data to pull in the updated data for the target field
' -----------------------------------------------------------------------------

    For i = 2 To UBound(arry_Match_Dest)
        If dict_LookupData.Exists(arry_Match_Dest(i)) Then
            arry_Target_Dest(i) = dict_LookupData.Item(arry_Match_Dest(i))
        End If
    Next i

    ' Output the values from the array
    rng_Target_Dest.Value2 = Application.Transpose(arry_Target_Dest)
    
    ' Empty the Dictionary
    dict_LookupData.RemoveAll

' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function

Function fx_Count_on_Single_Field(wsSource As Worksheet, wsDest As Worksheet, str_Source_TargetField As String, str_Source_MatchField As String, str_Dest_TargetField As String, str_Dest_MatchField As String, str_Criteria As String, Optional bol_CloseSourceWb As Boolean)

' Purpose: To import the summed / counted data from the Source to the Destination.
    ' 1) Add the data to be referenced into the dictionary
    ' 2) Look for the data to be updated for matches to the reference dictionary
    ' 3) Replace the existing data with the reference data

' Trigger: Called
' Updated: 5/5/2023

' Change Log:
'       6/29/2021:  Initial Creation, based on fx_Update_Single_Field
'       6/29/2021:  Added the 'And arry_Target_Source(i - 1) = str_Criteria' to ensure we were comparing the same Exception AND the same Account Number
'       11/16/2021: Added the strSumField
'       9/15/2022:  Split out the Count and Sum functions to simplify
'       5/5/2023:   Updated so that the int_LastRow_wsSource is a minimum of 2

' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Count_on_Single_Field( _
        wsSource:=ws5010a, wsDest:=wsPolicyExceptions_LL, _
        str_Source_TargetField:="Exception Name", _
        str_Source_MatchField:="14 Digit Account Number", _
        str_Dest_TargetField:="1 - Collateral (#)", _
        str_Dest_MatchField:="Account Number", _
        str_Criteria:="1 - Collateral & Other Support")

' LEGEND MANDATORY:
'   TBD:

' LEGEND OPTIONAL:
'   bolNotBlankSum: When the value in the terget isn't blank then do the Sum

' ***********************************************************************************************************************************

' -------------------------------
' Declare your wsSource variables
' -------------------------------
        
With wsSource
        
    ' Dim Integers
    
    Dim int_LastCol_wsSource As Long
        int_LastCol_wsSource = .Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim int_LastRow_wsSource As Long
        int_LastRow_wsSource = WorksheetFunction.Max( _
        .Cells(Rows.Count, "A").End(xlUp).Row, _
        .Cells(Rows.Count, "B").End(xlUp).Row, _
        .Cells(Rows.Count, "C").End(xlUp).Row)
        
        If int_LastRow_wsSource = 1 Then int_LastRow_wsSource = 2
        
    ' Dim "Ranges"
        
    Dim arry_Header_wsSource() As Variant
        arry_Header_wsSource = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsSource)))
        
    Dim col_Target_Source As Integer
        col_Target_Source = fx_Create_Headers(str_Source_TargetField, arry_Header_wsSource)

    Dim col_Match_Source As Integer
        col_Match_Source = fx_Create_Headers(str_Source_MatchField, arry_Header_wsSource)
        
    ' Dim Arrays
    
    Dim arry_Target_Source() As Variant
        arry_Target_Source = Application.Transpose(.Range(.Cells(1, col_Target_Source), .Cells(int_LastRow_wsSource, col_Target_Source)))
    
    Dim arry_Match_Source() As Variant
        arry_Match_Source = Application.Transpose(.Range(.Cells(1, col_Match_Source), .Cells(int_LastRow_wsSource, col_Match_Source)))
        
End With
        
' -----------------------------
' Declare your wsDest variables
' -----------------------------
        
With wsDest
        
    ' Dim wsDest Integers
    
    Dim int_LastCol_wsDest As Long
        int_LastCol_wsDest = wsDest.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim int_LastRow_wsDest As Long
        int_LastRow_wsDest = WorksheetFunction.Max( _
        wsDest.Cells(Rows.Count, "A").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "B").End(xlUp).Row, _
        wsDest.Cells(Rows.Count, "C").End(xlUp).Row)
 
    ' Dim wsDest "Ranges"
    
    Dim arry_Header_wsDest() As Variant
        arry_Header_wsDest = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol_wsDest)))
        
    Dim col_Target_Dest As Integer
        col_Target_Dest = fx_Create_Headers(str_Dest_TargetField, arry_Header_wsDest)

    Dim col_Match_Dest As Integer
        col_Match_Dest = fx_Create_Headers(str_Dest_MatchField, arry_Header_wsDest)
 
    ' Dim Ranges
        
    Dim rng_Target_Dest As Range
    Set rng_Target_Dest = .Range(.Cells(1, col_Target_Dest), .Cells(int_LastRow_wsDest, col_Target_Dest))
        
    ' Dim Arrays
    
    Dim arry_Target_Dest() As Variant
        arry_Target_Dest = Application.Transpose(.Range(.Cells(1, col_Target_Dest), .Cells(int_LastRow_wsDest, col_Target_Dest)))
    
    Dim arry_Match_Dest() As Variant
        arry_Match_Dest = Application.Transpose(.Range(.Cells(1, col_Match_Dest), .Cells(int_LastRow_wsDest, col_Match_Dest)))
 
End With
        
' ----------------------------
' Declare your Other variables
' ----------------------------
        
    ' Dim Dictionaries
    
    Dim dict_LookupData As Scripting.Dictionary
        Set dict_LookupData = New Scripting.Dictionary
        dict_LookupData.CompareMode = TextCompare
        
    ' Declare Loop Variables
    
    Dim i As Long
        
    Dim intCount As Long

' ----------------------------------------
' Fill the Dictionary with the Lookup Data
' ----------------------------------------
    
On Error Resume Next
       
        For i = 2 To UBound(arry_Target_Source)
            If arry_Target_Source(i) = str_Criteria Then
                If arry_Match_Source(i) = arry_Match_Source(i - 1) And arry_Target_Source(i - 1) = str_Criteria Then
                    intCount = intCount + 1
                Else
                    intCount = intCount + 1
                    dict_LookupData.Add Key:=arry_Match_Source(i), Item:=intCount
                    intCount = 0
                End If
            
                ' To handle the last record
                If i = UBound(arry_Target_Source) Then
                    intCount = intCount + 1
                    dict_LookupData.Add Key:=arry_Match_Source(i), Item:=intCount
                    intCount = 0
                End If
            
            End If
        
        Next i
        
On Error GoTo 0
    
' -----------------------------------------------------------------------------
' Loop through the Lookup Data to pull in the updated data for the target field
' -----------------------------------------------------------------------------

    For i = 2 To UBound(arry_Match_Dest)
        If dict_LookupData.Exists(arry_Match_Dest(i)) Then
            arry_Target_Dest(i) = dict_LookupData.Item(arry_Match_Dest(i))
        End If
    Next i

    ' Output the values from the array
    rng_Target_Dest.Value2 = Application.Transpose(arry_Target_Dest)
    
    ' Empty the Dictionary
    dict_LookupData.RemoveAll

' -------------------------
' Close the Source workbook
' -------------------------
    
    If wsSource.Parent.Name <> ThisWorkbook.Name And bol_CloseSourceWb = True Then
        wsSource.Parent.Close savechanges:=False
    End If

End Function
Function fx_Hide_Worksheets_For_Users()

' Purpose: To hide the extra worksheets that the "regular" users don't need to see.
' Trigger: Called
' Updated: 12/22/2020

' Change Log:
'       12/22/2020: Initial creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim ws As Worksheet

' --------------------------------
' Loop through the worksheet names
' --------------------------------

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard Review" And ws.Name <> "INSTRUCTIONS" Then
            ws.Visible = xlSheetHidden
        Else
            ws.Visible = xlSheetVisible
        End If
    Next ws

End Function

Function fx_Output_Unique_Values(rngSourceData As Range, rngDestLoc As Range)

' Purpose: To create a unique list of values based on the passed range.
' Trigger: Called
' Updated: 3/15/2021

' Use Example: _
    Call fx_Output_Unique_Values( _
        rngSourceData:=wsBO5356.Range("H2:H" & int_LastRow), _
        rngDestLoc:=wsBO5356.Range("J2"))

' Change Log:
'       5/5/2020: Intial Creation
'       6/2/2020: Updated the code to output an Array so it's easier to pull data from
'       3/15/2021: Updated to pass the rngSource and rngDest
'       3/15/2021: Updated the output to resolve the issue with passing an array to a static destination

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
           
    Dim strsUniqueValues As New Collection
    
    Dim strValue As Variant
    
    Dim arryTempData() As Variant
        arryTempData = rngSourceData
    
    Dim i As Long
           
' -----------
' Pass the values from the Temp array to the collection (to pull in only unique values)
' -----------

On Error Resume Next 'When a duplicate is found skip it, instead of erroring
    
    For Each strValue In arryTempData
        strsUniqueValues.Add strValue, strValue
    Next

On Error GoTo 0

' -----------
' Pass the array of unique values
' -----------

    Erase arryTempData
        ReDim arryTempData(1 To strsUniqueValues.Count)
    
    'Add values to a one dimensional array
    For i = 1 To strsUniqueValues.Count
        arryTempData(i) = strsUniqueValues(i)
    Next i
    
    'Output the values from the array to the destination
    rngDestLoc.Resize(strsUniqueValues.Count) = Application.Transpose(arryTempData)

End Function
Function fx_Get_Dictionary_Key(dictLookup As Dictionary, strItem As String) As String

' Purpose: To lookup the Key for a passed dictionary Item.
' Trigger: Called
' Updated: 4/28/2021

' Note: Only works with unique items, if multiples will output the Key for the first Item found.

' Use Example:
'    Me.cmb_ParentCode = fx_Get_Dictionary_Key(dict_ParentCodes, Me.lst_Customers.Value)

' Change Log:
'       4/28/2021: Intial Creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim varKey As Variant

' --------------------------------
' Add the values to the dictionary
' --------------------------------

    For Each varKey In dictLookup.Keys
        If dictLookup.Item(varKey) = strItem Then
            fx_Get_Dictionary_Key = CStr(varKey)
            Exit Function
        End If
    Next

End Function
Function fx_Steal_First_Row_Formating(ws As Worksheet, Optional intFirstRow As Long, Optional int_LastCol As Long, Optional int_LastRow As Long, Optional intSingleRow As Long)

' Purpose: To copy the formatting from the first row of data and apply to the rest of the data.
' Trigger: Called
' Updated: 1/19/2022

' Use Example: _
    Call fx_Steal_First_Row_Formating( _
        ws:=wsQCReview, _
        intFirstRow:=2, _
        int_LastRow:=int_LastRow, _
        int_LastCol:=int_LastCol)

' Use Example 2: _
    Call fx_Steal_First_Row_Formating( _
        ws:=wsData, _
        intSingleRow:=CurRow)

' Use Example 3: Call fx_Steal_First_Row_Formating(ws:=wsQCReview)

' Change Log:
'       5/17/2021:  Intial Creation
'       6/16/2021:  Added the 'Application.Goto' to reset the copy paste
'       12/6/2021:  Added the option to pass only a single row
'       12/8/2021:  Added the rngCur so the screen doesn't jump around
'       1/8/2022:   Updated some of the passed variables to be optional
'                   Defaulted intFirstRow to be 2 if not passed
'       1/19/2022:  Added the ws. qualifiers for LastRow and LastCol

' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    Dim rngCur As Range
    Set rngCur = ActiveCell

    'Declare Integers
    
    If intFirstRow = 0 Then intFirstRow = 2
    
    If int_LastRow = 0 Then
       int_LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    End If
    
    If int_LastCol = 0 Then
       int_LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    End If
    
    'Declare Ranges
    
With ws

    Dim rngFormat As Range
        Set rngFormat = .Range(.Cells(intFirstRow, 1), .Cells(intFirstRow, int_LastCol))
    
    Dim rngTarget As Range
        If intSingleRow <> 0 Then
            Set rngTarget = .Range(.Cells(intSingleRow, 1), .Cells(intSingleRow, int_LastCol))
        ElseIf int_LastRow <> 0 Then
            Set rngTarget = .Range(.Cells(intFirstRow + 1, 1), .Cells(int_LastRow, int_LastCol))
        Else
            MsgBox "There was no row passed to the Steal First Row function."
        End If

End With

' ---------------------------------------------------------------------------------------------------
' Copy the formatting from the first row of data (intFirstRow) to the remaining rows (thru int_LastRow)
' ---------------------------------------------------------------------------------------------------
    
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteFormats
    
    'Go back to where you were before the code
    Application.CutCopyMode = False
    Application.Goto Reference:=rngCur, Scroll:=False
    
End Function

Function fx_Set_Cell_A1_As_Active()

' Purpose: To loop through each visible worksheet and set Cell A1 as the active cell for each.
' Trigger: Called
' Updated: 5/18/2021

' Change Log:
'       5/18/2021: Intial Creation

' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    'Declare Worksheets
    
    Dim ws As Worksheet

' -----------
' Set A1 as the Active Cell
' -----------

    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Visible = xlSheetVisible Then
            Application.Goto Reference:=ws.Range("A1"), Scroll:=False
        End If
    
    Next ws

End Function
Function fx_Create_File_Dictionary(strFolderPath As String) As Dictionary

' Purpose: To create a dictionary of files for a given folder path.
' Trigger: Called
' Updated: 8/9/2021

' Use Example: _
    Set dict_Files = fx_Create_File_Dictionary(strFolderPath:="C:\U Drive\Support\Weekly Plan\")

' Change Log:
'       8/9/2021: Initial Creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare Objects

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objDirParent As Object
    Set objDirParent = objFSO.GetFolder(strFolderPath)
    
    Dim objFile As Object
    
    'Declare Dictionary
    
    Dim dict_File As Scripting.Dictionary
    Set dict_File = New Scripting.Dictionary

' ----------------------------------
' Load the Files into the dictionary
' ----------------------------------

On Error Resume Next

    For Each objFile In objDirParent.Files
        dict_File.Add Key:=objFile.Name, Item:=objFile.Path
    Next objFile
                        
On Error GoTo 0

' Output the results
Set fx_Create_File_Dictionary = dict_File

End Function
Function fx_Update_Default_Directory()

' Purpose: This function will reset the Drive and Directory to wherever ThisWorkbook is located.
' Trigger: Called Function
' Updated: 8/12/2021

' Change Log:
'       8/12/2021: Initial Creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
' ----------------------------------
' Set the current directory and path
' ----------------------------------
    
On Error Resume Next
    ChDrive ThisWorkbook.Path
        
        If objFSO.FolderExists(ThisWorkbook.Path & "\(Source Data)") = True Then
            ChDir ThisWorkbook.Path & "\(Source Data)"
        Else
            ChDir ThisWorkbook.Path
        End If
        
On Error GoTo 0

    'Release the Object
    Set objFSO = Nothing

End Function
Public Function fx_File_Exists(strFullPath As String) As Boolean

' Purpose: This function will determine if a file exists already.
' Trigger: Called
' Updated: 7/18/2022

' Use Example: _
    'bolTodayFileExists = fx_File_Exists(strTodayFileLoc)

' Change Log:
'       8/19/2021:  Initial Creation
'       7/18/2022:  Added the Use Example

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Dim Objects
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
' ---------------------------
' Check if passed file exists
' ---------------------------
    
    If objFSO.FileExists(strFullPath) = True Then fx_File_Exists = True

    'Release the Object
    Set objFSO = Nothing

End Function
Function fx_Delete_Unused_Data(ws_Target As Worksheet, str_Target_Field As String, str_Value_To_Delete As String, Optional bol_DeleteDataOnly As Boolean, Optional bol_DeleteValues_PassArray As Boolean)

' Purpose: To delete data from the passed worksheet where the "Value To Delete" is in the Target Field.
' Trigger: Called
' Updated: 6/29/2023
'
' Change Log:
'       9/15/2021:  Initial Creation
'       5/5/2023:   Added the 'Exit Function' if the int_LastRow is 1, to handle situations where there is no data (ex. Policy Exceptions - Paid Off Loans)
'                   Updated the delete to keep the first row of data for the formatting
'       6/12/2023:  Added the option to pass an array of values to be deleted, and the code around arryValuesToDelete
'       6/29/2023:  Added the option to delete just the values, not the rows
'       7/3/2023:   Updated to use the Target Field as the anchor for int_LastRow and added mroe detailed debugging for int_LastRow
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Delete_Unused_Data( _
        ws_Target:=wsSageworksRT_Dest, _
        str_Target_Field:="Line of Business", _
        str_Value_To_Delete:="Small Business")

' LEGEND OPTIONAL:
'   bol_DeleteDataOnly: Allows only the specific fields to be deleted, otherwise the entire row is deleted
'   bol_FilterPassArray: Allows an array of values to be passed instead of a single value for the filter



' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

With wsTarget

    ' Declare Cell References / "Ranges"
       
    Dim int_LastCol As Integer
        int_LastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol)))

    Dim col_Target As Integer
        col_Target = fx_Create_Headers(str_Target_Field, arry_Header)
       
    Dim int_LastRow As Long
        int_LastRow = .Cells(Rows.Count, col_Target).End(xlUp).Row

    If int_LastRow = 1 Then
        Debug.Print "The 'fx_Delete_Unused_Data' function didn't run on " & Now & Chr(10) & _
                    "This is becuase the int_LastRow variable would have been 1."
        Exit Function
    End If

    ' Declare Arrays
    
    Dim arryValuesToDelete() As String
    
    If bol_DeleteValues_PassArray = True Then
        arryValuesToDelete = Split(str_Value_To_Delete, ", ")
    End If
    
End If
    
' ----------------------
' Delete the Unused Data
' ----------------------
        
On Error Resume Next
        
With wsTarget
    
    ' Sort the data to make deleting MUCH faster
    .Range(.Cells(1, 1), .Cells(int_LastRow, int_LastCol)).Sort _
        Key1:=.Cells(1, col_Target), Order1:=xlAscending, Header:=xlYes
    
    If bol_DeleteDataOnly = True Then
        
        If bol_DeleteValues_PassArray = False Then
        
            .Range("A1").AutoFilter Field:=col_Target, Criteria1:=str_Value_To_Delete, Operator:=xlFilterValues
                .Range(.Cells(2, col_Target), .Cells(int_LastRow, col_Target)).SpecialCells(xlCellTypeVisible).ClearContents
                .Range("A1").AutoFilter Field:=col_Target
        Else
            .Range("A1").AutoFilter Field:=col_Target, Criteria1:=arryValuesToDelete, Operator:=xlFilterValues
                .Range(.Cells(2, col_Target), .Cells(int_LastRow, col_Target)).SpecialCells(xlCellTypeVisible).ClearContents
                .Range("A1").AutoFilter Field:=col_Target
        
        End If
        
        Exit Function ' Skip the deleting the rows
    
    End If
    
    ' Filter based on the values and then delete the filtered rows
    If bol_DeleteValues_PassArray = False Then
    
        .Range("A1").AutoFilter Field:=col_Target, Criteria1:=str_Value_To_Delete, Operator:=xlFilterValues
            .Range("A2:A3").SpecialCells(xlCellTypeVisible).EntireRow.ClearContents
            .Range("A3:A" & int_LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_Target
    Else
        .Range("A1").AutoFilter Field:=col_Target, Criteria1:=arryValuesToDelete, Operator:=xlFilterValues
            .Range("A2:A3").SpecialCells(xlCellTypeVisible).EntireRow.ClearContents
            .Range("A3:A" & int_LastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            .Range("A1").AutoFilter Field:=col_Target
    
    End If
    
End With

On Error GoTo 0

End Function
Function fx_Rename_Header_Field(ws As Worksheet, intHeaderRow As Long, str_Value_To_Update As String, str_Updated_Value As String)

' Purpose: To rename one of the fields in the header of the passed workbook.
' Trigger: Called
' Updated: 9/16/2021
'
' Change Log:
'       9/16/2021: Initial Creation
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Rename_Header_Field( _
        ws:=wsSageworksRT_Dest, _
        intHeaderRow:=8, _
        str_Value_To_Update:="RTB High (Branch)", _
        str_Updated_Value:="RTB High")



' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare Cell References
              
    Dim int_LastCol As Long
        int_LastCol = ws.Cells(intHeaderRow, Columns.Count).End(xlToLeft).Column

    'Declare "Ranges"
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(ws.Range(ws.Cells(intHeaderRow, 1), ws.Cells(intHeaderRow, int_LastCol)))

    Dim col_Target As Long
        col_Target = fx_Create_Headers(str_Value_To_Update, arry_Header)
    
' -----------------------
' Update the Header Field
' -----------------------
        
On Error Resume Next
    
    ws.Cells(intHeaderRow, col_Target).Value2 = str_Updated_Value
    '.Cells(2, .Rows(2).Find("BookBalance").Column).Value = "Outstanding"
    
On Error GoTo 0

'       9/16/2021: To be added
'        > If it's missing then go to errhandler and output a message box
'            > Use the message box to determine if I continue or abort

End Function
Function fx_Rename_Header_Fields(ws_Target As Worksheet, intHeaderRow As Long)

' Purpose: To rename multiples fields in the header of the passed worksheet, based on the conversion table in wsLists.
' Trigger: Called
' Updated: 11/29/2021
'
' Change Log:
'       11/29/2021: Initial Creation, based on fx_Rename_Header_Field
'                   Updated to convert to using the arry_Header for the conversions
'                   If the applicable pair has strikethrough formatting ignore it
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
'   Call fx_Rename_Header_Fields(ws_Target:=ThisWorkbook.Sheets("TEST"), intHeaderRow:=1)

' ***********************************************************************************************************************************
    
' Abort if the wsLists doesn't exist in the Workbook
If fx_Sheet_Exists(ThisWorkbook.Name, "Lists") = False Then
    MsgBox "The wsLists sheet doesn't exists, fx_Rename_Header_Fields cannot run"
    Exit Function
End If

' -----------------
' Declare Variables
' -----------------

    'Declare Worksheets
    
    Dim wsLists As Worksheet
    Set wsLists = ThisWorkbook.Sheets("LISTS")
    
    'Declare Integers
    
    Dim int_LastCol As Long
        int_LastCol = ws_Target.Cells(intHeaderRow, Columns.Count).End(xlToLeft).Column
    
    Dim int_LastCol_wsLists As Long
        int_LastCol_wsLists = wsLists.Cells(1, Columns.Count).End(xlToLeft).Column
    
    Dim int_LastRow_wsLists As Long
            
    'Declare Source Cell References

    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(ws_Target.Range(ws_Target.Cells(intHeaderRow, 1), ws_Target.Cells(intHeaderRow, int_LastCol)))

    'Declare wsList Cell References
    
    Dim arry_Header_wsLists() As Variant
        arry_Header_wsLists = Application.Transpose(wsLists.Range(wsLists.Cells(1, 1), wsLists.Cells(1, int_LastCol_wsLists)))
    
    Dim col_OriginalField As Long
        col_OriginalField = fx_Create_Headers("Original Field", arry_Header_wsLists)
    
    Dim col_UpdatedField As Long
        col_UpdatedField = fx_Create_Headers("Updated Field", arry_Header_wsLists)
    
    'Declare Loop Variables
    
    Dim intHeaderCounter As Long
    
    Dim intConversionCounter As Long
    
    Dim i As Long
    
    'Declare Dictionaries
    
    Dim dictFieldConversion As New Scripting.Dictionary
        dictFieldConversion.CompareMode = TextCompare
    
' ------------------------------------
' Fill the Field Conversion Dictionary
' ------------------------------------
    
    int_LastRow_wsLists = wsLists.Cells(Rows.Count, col_OriginalField).End(xlUp).Row

    For i = 2 To int_LastRow_wsLists
        If wsLists.Cells(i, col_OriginalField).Font.Strikethrough = False Then
            dictFieldConversion.Add Key:=wsLists.Cells(i, col_OriginalField).Value2, Item:=wsLists.Cells(i, col_UpdatedField).Value2
        End If
    Next i
    
' -----------------------
' Update the Header Field
' -----------------------
        
On Error Resume Next
    
    For intHeaderCounter = 1 To int_LastCol
    
        If dictFieldConversion.Exists(arry_Header(intHeaderCounter, intHeaderRow)) Then
            arry_Header(intHeaderCounter, intHeaderRow) = dictFieldConversion.Item(arry_Header(intHeaderCounter, intHeaderRow))
        End If
    
    Next intHeaderCounter

On Error GoTo 0

' -----------------------
' Output the Header Field
' -----------------------

    ws_Target.Range(ws_Target.Cells(intHeaderRow, 1), ws_Target.Cells(intHeaderRow, int_LastCol)) = Application.Transpose(arry_Header)

End Function
Public Function fx_Return_Column_Letter(intColumnNum As Long) As String

' Purpose: This function will return the letter of the passed column.
' Trigger: Called
' Updated: 9/16/2021

' Change Log:
'       9/16/2021: Initial Creation

' ***********************************************************************************************************************************
    
    fx_Return_Column_Letter = Split(Cells(1, intColumnNum).Address, "$")(1)

End Function
Public Function fx_Return_Column_Number(strColumnLetter As String) As Long

' Purpose: This function will return the number of the passed column based on the letter.
' Trigger: Called
' Updated: 11/9/2023

' Change Log:
'       11/9/2023: Initial Creation

' ***********************************************************************************************************************************
    
    fx_Return_Column_Number = Range(strColumnLetter & "1").Column

End Function
Function fx_Clear_Old_Data(ws As Worksheet, Optional bolDeleteHeaderRow As Boolean, Optional bolClearFormatting As Boolean, Optional bolKeepFirstDataRowFormatting As Boolean)

' Purpose: To clear the existing data as the first step of an import process.
' Trigger: Called
' Updated: 8/2/2022
'
' Change Log:
'       10/5/2021:  Intial Creation
'       10/20/2021: Updated to include the firstrow and allow formatting to be cleared.
'       1/8/2022:   Updated so that the first row of formatting gets retained
'       1/31/2022:  Allow the first row of data to be retained, to keep formulas intact
'                   Made 'bolClearFormatting' and 'bolDeleteHeader' Optional
'       8/2/2022:   Added the line of code to unmerge the cells
'                   Updated int_LastCol to use SpecialCells(xlCellTypeLastCell)
'                   Updated to not do 'intFirstRow + 1' in the ClearFormatting, was fighting with bolKeepFirstRow
'                   Created the intFirstRowFormatting variable and renamed intFirstRow to intFirstRowData
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Clear_Old_Data( _
        ws:=wsSageworksRT_Dest, _
        bolDeleteHeaderRow:=False, _
        bolKeepFirstDataRowFormatting:=True, _
        bolClearFormatting:=True)

' LEGEND:
'   bolKeepFirstRowFormatting: If this is true then I will keep the formatting for the first row of data
'   bolClearFormatting: If this is true I will remove all formatting


' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max( _
            ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
            ws.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row, _
            ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row)
            
        If int_LastRow = 1 Then int_LastRow = 2
            
    Dim intFirstRowData As Long ' First row of data to be removed
        intFirstRowData = 2
        
        If bolDeleteHeaderRow = True Then
            intFirstRowData = 1
        End If
    
    Dim intFirstRowFormatting As Long
        intFirstRowFormatting = intFirstRowData
    
        If bolKeepFirstDataRowFormatting = True Then
            intFirstRowFormatting = 3
        End If
    
    Dim int_LastCol As Long
        int_LastCol = WorksheetFunction.Max( _
        ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column, _
        ws.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Column, _
        ws.Rows(1).Find("").Column - 1)

' -------------
' Wipe Old Data
' -------------
   
    ws.Range(ws.Cells(intFirstRowData, 1), ws.Cells(int_LastRow, int_LastCol)).UnMerge
    
    ws.Range(ws.Cells(intFirstRowData, 1), ws.Cells(int_LastRow, int_LastCol)).ClearContents
    
    If bolClearFormatting = True Then
        ws.Range(ws.Cells(intFirstRowFormatting, 1), ws.Cells(int_LastRow, int_LastCol)).ClearFormats
    End If

End Function
Function fx_Update_Named_Range(strNamedRangeName As String)

' Purpose: To update the passed Named Range a change in the Change Log.
' Trigger: Called
' Updated: 3/6/2022
'
' Change Log:
'       12/8/2021:  Intial Creation
'       3/4/2022:   Added Error Handling for the int_LastRow and int_LastCol to handle if all of the empty rows/cols are hidden
'       3/6/2022:   Replaced the int_LastRow and int_LastCol w/ functions
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Update_Named_Range("ChangeLog_Data")



' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim str_WsName As String ' Get the name of the worksheet where the Named Range is
        str_WsName = ThisWorkbook.Names(strNamedRangeName).RefersToRange.Parent.Name
    
    Dim wsNamedRange As Worksheet
    Set wsNamedRange = ThisWorkbook.Sheets(str_WsName)
    
    Dim int_LastRow As Long
        int_LastRow = fx_Find_LastRow(wsNamedRange)
        
    Dim int_LastCol As Integer
        int_LastCol = fx_Find_LastColumn(wsNamedRange)

' ----------------------
' Update the Named Range
' ----------------------

   ThisWorkbook.Names(strNamedRangeName).RefersToR1C1 = wsNamedRange.Range(wsNamedRange.Cells(1, 1), wsNamedRange.Cells(int_LastRow, int_LastCol))

End Function

Public Function fx_Load_ListBox_From_wsLists(str_TargetFieldName As String, Optional lstTarget As MSForms.ListBox, Optional cmbTarget As MSForms.ComboBox, Optional bolUniqueOnly As Boolean)

' Purpose: This function will load values from wsLists into the applicable ListBox.
' Trigger: Called
' Updated: 2/10/2022
'
' Change Log:
'       1/15/2022:  Initial Creation
'       1/18/2022:  Updated to load data into a ComboBox
'                   Removed the ufUserFormName, it was redundant as it's needed for lstTarget or cmbTarget
'       2/10/2022:  Added the code for bolUniqueOnly to allow only the unique values to be passed
'                   Added the LoadForms code to let me skip the new steps
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call myFunctions.fx_Load_ListBox_From_wsLists(lstTarget:=uf_TEST.lst_TEST, str_TargetFieldName:="Role")


' ***********************************************************************************************************************************
    
' -----------------
' Declare Variables
' -----------------
    
    'Declare Worksheets
    
    If Evaluate("ISREF(" & "'Lists'" & "!A1)") = True Then
        Dim wsLists As Worksheet
        Set wsLists = ThisWorkbook.Sheets("Lists")
    Else
        MsgBox "wsLists doesn't exist and therefore fx_Load_ListBox_From_wsLists can't run."
    End If

    'Declare Cell References
    
    Dim int_LastCol As Integer
        int_LastCol = wsLists.Cells(1, Columns.Count).End(xlToLeft).Column
        
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(wsLists.Range(wsLists.Cells(1, 1), wsLists.Cells(1, int_LastCol)))

    Dim col_TargetField As Long
        col_TargetField = fx_Create_Headers(str_TargetFieldName, arry_Header)

    Dim int_LastRow As Long
        int_LastRow = wsLists.Cells(Rows.Count, col_TargetField).End(xlUp).Row

    'Declare Arrays
    
    Dim arryTargetField() As Variant
        arryTargetField = wsLists.Range(wsLists.Cells(2, col_TargetField), wsLists.Cells(int_LastRow, col_TargetField))
    
' ------------------------------------------------
' Determine if I need to output Unique Values only
' ------------------------------------------------
    
    If bolUniqueOnly = True Then
        GoTo OutputUniqueValues
    Else
        GoTo LoadForms
    End If
    
' ----------------------------------
' Declare Variables for UniqueValues
' ----------------------------------
           
OutputUniqueValues:
           
    Dim strsUniqueValues As New Collection
    
    Dim strValue As Variant
    
    Dim arryTempData() As Variant
        arryTempData = arryTargetField
    
    Dim i As Long
           
' -------------------------------------------------------------------------------------
' Pass the values from the Temp array to the collection (to pull in only unique values)
' -------------------------------------------------------------------------------------

On Error Resume Next 'When a duplicate is found skip it, instead of erroring
    
    For Each strValue In arryTempData
        strsUniqueValues.Add strValue, strValue
    Next

On Error GoTo 0

' -------------------------------
' Pass the array of unique values
' -------------------------------

    Erase arryTempData
        ReDim arryTempData(1 To strsUniqueValues.Count)
    
    'Add values to a one dimensional array
    For i = 1 To strsUniqueValues.Count
        arryTempData(i) = strsUniqueValues(i)
    Next i

    'Save the temp data into the arryTargetField
    
    Erase arryTargetField
    
    arryTargetField = arryTempData

' ------------------------------------
' Load the UserForm ListBox / ComboBox
' ------------------------------------
    
LoadForms:
    
    If Not lstTarget Is Nothing Then
        lstTarget.List = arryTargetField
    ElseIf Not cmbTarget Is Nothing Then
        cmbTarget.List = arryTargetField
    Else
        MsgBox "There was no Optional paramater past for lstTarget or cmbTarget, so there is nothing to update"
    End If

End Function
Function fx_Create_Dynamic_Lookup_List(wsDataSource As Worksheet, str_Dynamic_Lookup_Value As String, col_Dynamic_Lookup_Field As Long, Optional col_Criteria_Field As Long, Optional str_Criteria_Match_Value As Variant, Optional col_Target_Field As Long) As Variant

' Purpose: To create the dynamic list of values to be used in the ListBox, based on a change to the cmb_Dynamic_Borrower_Lookup.
' Trigger: Start typing in the Dynamic_Borrower_Lookup combo box (cmb_Dynamic_Borrower_Lookup_Change)
' Updated: 10/7/2022
'
' Change Log:
'       11/21/2021: Intial Creation for the PAR Agenda, taken from Sageworks Validation Dashboard code
'       1/18/2022:  Updated and converted to a function
'       10/6/2022:  Updated the naming of the fields, and handled the error if there was no LOB Lookup value
'       10/7/2022:  Updated to allow a seperate Dynamic Lookup and Target field
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
'    Dim arryBorrowersTemp As Variant
'       arryBorrowersTemp = myFunctions.fx_Create_Dynamic_Lookup_List( _
        wsDataSource:=wsData, _
        col_Dynamic_Lookup_Field:=col_Borrower, _
        str_Dynamic_Lookup_Value:=Me.cmb_Dynamic_Borrower.Value, _
        col_Criteria_Field:=col_LOBUpdated, _
        str_Criteria_Match_Value:=Me.lst_LOB.Value)

'    Me.lst_Borrowers.List = arryTargetValuesTemp

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare Strings
    Dim strLookupValue As String
    Dim strCriteriaValue As String
    Dim str_TargetValue As String
    
    'Declare Cell References
    Dim intSourceRow As Long: intSourceRow = 2
    Dim intArryRow As Long: intArryRow = 0
    
    'Declare Arrays
    Dim arryTargetValues As Variant
        ReDim arryTargetValues(1 To 99999)

    If IsMissing(str_Criteria_Match_Value) Then str_Criteria_Match_Value = ""

' -------------------------------------
' Add the borrowers to the lookup Array
' -------------------------------------
       
    With wsDataSource
            
        Do While .Cells(intSourceRow, col_Dynamic_Lookup_Field).Value2 <> ""
            
            ' Set the loop variables
            strLookupValue = .Cells(intSourceRow, col_Dynamic_Lookup_Field).Value2
            If col_Criteria_Field <> 0 Then strCriteriaValue = .Cells(intSourceRow, col_Criteria_Field).Value2
            If col_Target_Field <> 0 Then str_TargetValue = .Cells(intSourceRow, col_Target_Field).Value2
                
                'If the data matches add to the array
                If InStr(1, strLookupValue, str_Dynamic_Lookup_Value, vbTextCompare) Then
                    If col_Criteria_Field = 0 Or str_Criteria_Match_Value = strCriteriaValue Then
                        intArryRow = intArryRow + 1
                        If str_TargetValue <> "" Then arryTargetValues(intArryRow) = str_TargetValue Else arryTargetValues(intArryRow) = strLookupValue
                    End If
                End If
            
            intSourceRow = intSourceRow + 1
        Loop
    End With

    If intArryRow > 0 Then ' If nothing was passed, don't redim
        ReDim Preserve arryTargetValues(1 To intArryRow)
    End If
    
    'Output the results
    fx_Create_Dynamic_Lookup_List = arryTargetValues

End Function
Public Function fx_Copy_Data_from_Clipboard_to_String() As String

' Purpose: To copy the content from the clipboard to a string.
' Trigger: Called
' Updated: 1/22/2022

' Change Log:
'       1/22/2022:  Initial Creation

' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------

    Dim objDataObj As New MSForms.DataObject
        objDataObj.GetFromClipboard

On Error Resume Next
    Dim strClipboardContent As String
        strClipboardContent = objDataObj.GetText
On Error GoTo 0

    fx_Copy_Data_from_Clipboard_to_String = strClipboardContent

    'objDataObj = Empty

End Function
Public Function fx_Remove_Leading_and_Trailing_Spaces(ws_Target As Worksheet)

' Purpose: To remove the leading and trailing spaces in all of the cells in the passed worksheet.
' Trigger: Called
' Updated: 1/22/2022
'
' Change Log:
'       1/22/2022:  Initial Creation
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
'    Call fx_Remove_Leading_and_Trailing_Spaces(ws_Target:=wsInfoGrid)



' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim rngTarget As Range
    Set rngTarget = ws_Target.UsedRange

    Dim arryData() As Variant
        arryData = rngTarget

    Dim x As Long
    
    Dim y As Long

' ------------------------------------
' Remove the leading / trailing spaces
' ------------------------------------

    For x = 1 To UBound(arryData, 1)
        For y = 1 To UBound(arryData, 2)
            arryData(x, y) = Trim(arryData(x, y))
        Next y
    Next x
    
    rngTarget = arryData 'Save back the updated data

End Function
Public Function fx_Convert_String_Amounts_to_Numeric(strAmounttoConvert As String) As Variant

' Purpose: To convert various ways to type a number (ex. $5MM, $5k, etc) to an actual numeric amount.
' Trigger: Called
' Updated: 3/4/2022
'
' Change Log:
'       1/22/2022:  Initial Creation
'       1/23/2022:  Added the conversion for Billions
'                   Removed the 'strToReplace', don't need as I replaced w/ the loop
'                   Updated to include a . if it's in there
'       1/26/2022:  Simplified to remove the "MM" and other duplicative references {M is in MM}
'       1/27/2022:  Converted to a Variant type and allowed for strings to be passed (in case the amount is actually text)
'       2/1/2022:   Added the code to abort if the string is to long (aka they are passing a bunch of text)
'       3/4/2022:   Updated so that M is Thousands and MM is Millions
'                   Added bolNegativeAmount to determine if the amount is negative
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
'    Call fx_Convert_String_Amounts_to_Numeric(strAmounttoConvert:="5MM")


' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim intConvertedAmount As Double
        
    Dim strAbbreviation As String
    
    Dim strNumericOnly As String
    
    Dim bolNegativeAmount As Boolean
    
    'Dim Loop Variables
    
    Dim i As Long
    
' -------------------------------------------------
' Determine if the amount has a bunch of text in it
' -------------------------------------------------

    If Len(strAmounttoConvert) > 20 Then
        fx_Convert_String_Amounts_to_Numeric = strAmounttoConvert
        Exit Function
    End If
    
' -------------------------------------------------------------------
' Determine if the amount is in Millions or Thousands or Whole Dollar
' -------------------------------------------------------------------
    
    If InStr(1, strAmounttoConvert, "B", vbTextCompare) Then
        strAbbreviation = "Billion"
    ElseIf InStr(1, strAmounttoConvert, "MM", vbTextCompare) Then
        strAbbreviation = "Million"
    ElseIf InStr(1, strAmounttoConvert, "K", vbTextCompare) Then
        strAbbreviation = "Thousand"
    ElseIf InStr(1, strAmounttoConvert, "M", vbTextCompare) Then
        strAbbreviation = "Thousand"
    End If
        
' -------------------------------------
' Determine if the amount is a negative
' -------------------------------------
        
    If Left(strAmounttoConvert, 1) = "(" Then
        bolNegativeAmount = True
    End If
        
' ---------------------------------
' Do the conversion to whole dollar
' ---------------------------------
        
    For i = 1 To Len(strAmounttoConvert)
        If Mid(strAmounttoConvert, i, 1) >= "0" And Mid(strAmounttoConvert, i, 1) <= "9" Then
            strNumericOnly = strNumericOnly + Mid(strAmounttoConvert, i, 1)
        ElseIf Mid(strAmounttoConvert, i, 1) = "." Then
            strNumericOnly = strNumericOnly + Mid(strAmounttoConvert, i, 1)
        End If
    Next
    
    If strNumericOnly <> "" Then
        intConvertedAmount = val(strNumericOnly) 'Convert all of the #s in the string to a value
        If bolNegativeAmount = True Then intConvertedAmount = -intConvertedAmount
    Else
        fx_Convert_String_Amounts_to_Numeric = strAmounttoConvert
        Exit Function
    End If
    
    If strAbbreviation = "Billion" Then
        intConvertedAmount = intConvertedAmount * 10 ^ 9
    ElseIf strAbbreviation = "Million" Then
        intConvertedAmount = intConvertedAmount * 10 ^ 6
    ElseIf strAbbreviation = "Thousand" Then
        intConvertedAmount = intConvertedAmount * 10 ^ 3
    Else
        intConvertedAmount = intConvertedAmount
    End If

    fx_Convert_String_Amounts_to_Numeric = Format(intConvertedAmount, "$#,##0")

End Function

Public Function fx_Sort_Collection(collTarget As Collection, Optional intFirstRecordNum As Long, Optional intLastRecordNum As Long) As Collection
  
' Purpose: This function will load values from wsLists into the applicable ListBox.
' Trigger: Called
' Updated: 12/8/2020
'
' Change Log:
'       12/8/2020:   Initial Creation, in the Sageworks Validation Dashboard, based on code from Axcel
'       2/10/2022:   Added as a function to myFunctions and updated all the wording
'                    Switched the FirstRecordNum and LastRecordNum to be optional
'
' -----------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    'Set coll_SortedPMs = fx_Sort_Collection(coll:=coll_UniquePMs, intFirstRecordNum:=1, intLastRecordNum:=coll_UniquePMs.count)



' ***********************************************************************************************************************************
  
' Assign the Optional variables if they weren't passed
If intFirstRecordNum = 0 Then intFirstRecordNum = 1
If intLastRecordNum = 0 Then intLastRecordNum = collTarget.Count
  
' -----------------
' Declare Variables
' -----------------
  
    'Declare Variants
    
    Dim vCentreVal As Variant
    Dim vTemp As Variant
    
    'Declare Integers
    
    Dim intTempLow As Long
        intTempLow = intFirstRecordNum
    
    Dim intTempHi As Long
        intTempHi = intLastRecordNum

' -------------------------------
' Sort the data in the collection
' -------------------------------
    
    vCentreVal = collTarget((intFirstRecordNum + intLastRecordNum) \ 2)
    Do While intTempLow <= intTempHi
    
      Do While collTarget(intTempLow) < vCentreVal And intTempLow < intLastRecordNum
        intTempLow = intTempLow + 1
      Loop
    
    Do While vCentreVal < collTarget(intTempHi) And intTempHi > intFirstRecordNum
      intTempHi = intTempHi - 1
    Loop
    
    If intTempLow <= intTempHi Then
    
      'Swap values
      vTemp = collTarget(intTempLow)
      
      collTarget.Add collTarget(intTempHi), After:=intTempLow
      collTarget.Remove intTempLow
      
      collTarget.Add vTemp, Before:=intTempHi
      collTarget.Remove intTempHi + 1
      
      'Move to next positions
      intTempLow = intTempLow + 1
      intTempHi = intTempHi - 1
      
    End If
    
  Loop
  
  If intFirstRecordNum < intTempHi Then fx_Sort_Collection collTarget, intFirstRecordNum, intTempHi
  If intTempLow < intLastRecordNum Then fx_Sort_Collection collTarget, intTempLow, intLastRecordNum
  
  Set fx_Sort_Collection = collTarget
  
End Function
Function fx_Convert_to_Values(ws_Target As Worksheet, str_TargetField_Name As String, Optional int_FirstRowofData As Long)
    
' Purpose: To convert the passed range from text to values.
' Trigger: Called
' Updated: 7/3/2023

' Change Log:
'       7/3/2023:   Initial Creation, based on 'u_Convert_to_Values' from my To Do
'                   Overhauled to be more dynamic / resiliant

' ***********************************************************************************************************************************

' USE EXAMPLE: _
    Call fx_Convert_to_Values(ws_Target:=ThisWorkbook.Worksheets("Data"), str_TargetField_Name:="Account Number / Loan Number")

' LEGEND MANDATORY:
'   ws_Target: The worksheet where the data to be converted resides.
'   str_TargetField_Name: The name of the field where the target data to be convereted resides.

' LEGEND OPTIONAL:
'   int_FirstRowofData: If passed then use this as the int_FirstRow instead of 2.

' Note:
'       7/3/2023: Using an Array to load / read the data took 1/2 the time as updating the range directlym

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

With ws_Target

    ' Declare "Ranges" / Cell References
    
    Dim int_LastCol As Integer
        int_LastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
        If int_LastCol = 1 Then int_LastCol = 2
    
    Dim arry_Header() As Variant
        arry_Header = Application.Transpose(.Range(.Cells(1, 1), .Cells(1, int_LastCol)))

    Dim col_Target As Integer
        col_Target = fx_Create_Headers(str_TargetField_Name, arry_Header)
           
    Dim int_FirstRow As Long
        If int_FirstRowofData > 0 Then
            int_FirstRow = int_FirstRowofData
        Else
            int_FirstRow = 2
        End If
        
    Dim int_LastRow As Long
        int_LastRow = .Cells(Rows.Count, col_Target).End(xlUp).Row
        If int_LastRow = 1 Then int_LastRow = 2
    
    ' Declare Loop Variables
    
    Dim rng_TargetData As Range
    Set rng_TargetData = .Range(.Cells(int_FirstRow, col_Target), .Cells(int_LastRow, col_Target))
    
    Dim arry_TargetData() As Variant
        arry_TargetData = WorksheetFunction.Transpose(rng_TargetData.Value)
    
    Dim i As Long
    
End With

' -----------------------------------
' Attempt entire selection conversion
' -----------------------------------

On Error GoTo IndividualProcess

    rng_TargetData.NumberFormat = "0"
    rng_TargetData.Value = rng_TargetData.Value
    
Exit Function
    
' ---------------------------------------
' Do conversion at invidiual record level
' ---------------------------------------
    
IndividualProcess:

On Error Resume Next

    Debug.Print "There was an error in the 'fx_Convert_to_Values' function at " & Now & Chr(10) _
                & "As a result the conversion had to be done at the individual record level."
    
    For i = LBound(arry_TargetData) To UBound(arry_TargetData)
        If IsNumeric(arry_TargetData(i)) = True Then
            rng_TargetData(i, 1).Value = val(arry_TargetData(i))
        End If
    Next i
    
End Function
