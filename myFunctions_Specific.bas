Attribute VB_Name = "myFunctions_Specific"
Option Explicit
Function fx_Create_Unique_List(rngListValues As Range)

' Purpose: To create a unique list of values based on the passed range.
' Trigger: Called
' Updated: 5/5/2020

' Change Log:
'          5/5/2020: Intial Creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
           
    Dim strsUniqueValues As New Collection
    
    Dim strValue As Variant
    
    Dim arryTempData()
           
' ------------------------------------------------------------------------------------
' Copy in the data selected for rngListValues into the array, then into the collection
' ------------------------------------------------------------------------------------

    arryTempData = Application.Transpose(rngListValues)

On Error Resume Next 'When a duplicate is found skip it, instead of erroring
    
    For Each strValue In arryTempData
        strsUniqueValues.Add strValue, strValue
    Next

On Error GoTo 0

' ------------------------------------
' Pass the collection of unique values
' ------------------------------------

Set fx_Create_Unique_List = strsUniqueValues
    Debug.Print "Unique Records found: " & Format(fx_Create_Unique_List.Count, "0,0")

End Function
Function fx_Create_Unique_List_As_Array(rngListValues As Range)

' Purpose: To create a unique list of values (in an array) based on the passed range.
' Trigger: Called
' Updated: 11.29.2021

' Change Log:
'       5.5.2020: Intial Creation
'       6.2.2020: Updated the code to output an Array so it's easier to pull data from
'       11.29.2021: Cleaned up the code

' Use Example: _
'   arryUniqueBorrowers = fx_Create_Unique_List_As_Array(wsData.Range(wsData.Cells(1, col_Borrower), wsData.Cells(int_LastRow_wsData, col_Borrower)))

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
           
    Dim strsUniqueValues As New Collection
    
    Dim strValue As Variant
    
    Dim arryTempData()
    
    Dim i As Long
           
' ------------------------------------------------------------------------------------
' Copy in the data selected for rngListValues into the array, then into the collection
' ------------------------------------------------------------------------------------

    arryTempData = rngListValues

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
    
    fx_Create_Unique_List_As_Array = arryTempData

End Function
Function fx_Output_Summed_Values_By_Unique_Record_As_Array(aryFullData, intColUnique As Long, intColToBeSummed As Long)

' Purpose: To create a unique list of values based on the passed range.
' Trigger: Called
' Updated: 6/2/2020

' Change Log:
'          6/2/2020: Intial Creation

' ***********************************************************************************************************************************

' -----------
' Declare your variables
' -----------
           
    Dim strsUniqueValues As New Collection
    
    Dim strValue As Variant
    
    Dim x As Long
    
    Dim z As Long
    
    Dim i As Long
           
    Dim intSummed As Double
    
    Dim aryTemp
    
    Dim aryUniqueOnly
        aryUniqueOnly = Application.WorksheetFunction.Index(aryFullData, 0, intColUnique)
           
' -----------
' Copy in just the unique values from the intColUnique into the collection
' -----------

On Error Resume Next 'When a duplicate is found skip it, instead of erroring
    
    For Each strValue In aryUniqueOnly
        strsUniqueValues.Add strValue, strValue
    Next

On Error GoTo 0

' -----------
' Pass the collection of unique values to AryTemp
' -----------

    ReDim aryTemp(1 To strsUniqueValues.Count, 1 To 2)

    For i = 1 To strsUniqueValues.Count
        aryTemp(i, 1) = strsUniqueValues(i)
    Next i
    
' -----------
' Sum for each customer
' -----------
    
     For x = 1 To UBound(aryTemp)
    
        For z = 1 To UBound(aryFullData)
           
            If aryTemp(x, 1) = aryFullData(z, intColUnique) Then
                intSummed = intSummed + aryFullData(z, intColToBeSummed)
            End If

        Next z
        
        aryTemp(x, 2) = intSummed 'Output the total, then clear for next
            intSummed = Empty
        
    Next x
    
' -----------
' Pass the results to the function
' -----------

    fx_Output_Summed_Values_By_Unique_Record_As_Array = aryTemp

End Function

Function fx_Return_Quarter(dtInput As Date)

' Purpose: To output the quarter for the passed date.
' Trigger: Called
' Updated: 11/20/2020

' Change Log:
'          11/20/2020: Intial Creation
    
' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim intQuarter As Long
        intQuarter = DatePart("q", dtInput)

    Dim intYear As Long
        intYear = DatePart("yyyy", dtInput)

    Dim dtOutputQuarter As Date
    
' -----------
' Output the given quarter
' -----------
    
    If intQuarter = 1 Then
        dtOutputQuarter = WorksheetFunction.EoMonth(DateSerial(intYear, 3, 1), 0)
    ElseIf intQuarter = 2 Then
        dtOutputQuarter = WorksheetFunction.EoMonth(DateSerial(intYear, 6, 1), 0)
    ElseIf intQuarter = 3 Then
        dtOutputQuarter = WorksheetFunction.EoMonth(DateSerial(intYear, 9, 1), 0)
    ElseIf intQuarter = 4 Then
        dtOutputQuarter = WorksheetFunction.EoMonth(DateSerial(intYear, 12, 1), 0)
    End If
    
    fx_Return_Quarter = dtOutputQuarter

End Function
Function fx_Array_Contains_Value(arryToSearch As Variant, valToFind As Variant) As Boolean

' Purpose: To determine if the given value is present in the array being searched.
' Trigger:alled
' Updated: 11/25/2020

' Change Log:
'          11/25/2020: Intial Creation
    
' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim bolValueFound As Boolean

    Dim i As Long

' -----------
' Check for the value
' -----------
    
    For i = LBound(arryToSearch) To UBound(arryToSearch)
        If arryToSearch(i) = valToFind Then
            bolValueFound = True
            Exit For
        End If
    Next i
    
    fx_Array_Contains_Value = bolValueFound

End Function
Function fx_Reverse_Given_Name(strNametoReverse As String)

' Purpose: This function splits and reverses a name from LAST, FIRST to FIRST LAST
' Trigger: Called
' Updated: 12/8/2020

' Change Log:
'       12/8/2020: Initial creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    If InStr(strNametoReverse, ",") = False Then 'If they have a unique name then abort
        'fx_Reverse_Given_Name = Replace(strNametoReverse, ".", "")
        fx_Reverse_Given_Name = strNametoReverse
        Exit Function
    End If

Dim str_First_Name As String
    str_First_Name = Right(strNametoReverse, Len(strNametoReverse) - InStrRev(strNametoReverse, ",") - 1)

Dim str_Last_Name As String
    str_Last_Name = Left(strNametoReverse, InStrRev(strNametoReverse, ",") - 1)

Dim str_Full_Name As String
    str_Full_Name = str_First_Name & " " & str_Last_Name

' -----------
' Output the name
' -----------

fx_Reverse_Given_Name = Replace(str_Full_Name, ".", "") 'Output the new user name, removes any periods after a middle initial

End Function
Public Function fx_Name_Reverse()

' Purpose: This function splits and reverses a user name from LAST, FIRST to FIRST LAST
' Trigger: Called
' Updated: 4/1/2020

' Change Log:
'       11/21/2021: Fixed an issue with people with middle names (ex."Elias, Richard J.")
'       4/1/2020: Fixed an issue with people with unique name formatting (First Last) that was breaking due to the missing comma.

' ***********************************************************************************************************************************

    Dim str_User_Name As String
        str_User_Name = Application.UserName
    
    If InStr(str_User_Name, ",") = False Then 'If they have a unique name then abort
        fx_Name_Reverse = Replace(str_User_Name, ".", "")
        Exit Function
    End If
    
    Dim str_First_Name As String
        str_First_Name = Right(str_User_Name, Len(str_User_Name) - InStrRev(str_User_Name, ",") - 1)
    
    Dim str_Last_Name As String
        str_Last_Name = Left(str_User_Name, InStrRev(str_User_Name, ",") - 1)
    
    Dim str_Full_Name As String
        str_Full_Name = str_First_Name & " " & str_Last_Name
    
    Debug.Print str_Full_Name
    
    fx_Name_Reverse = Replace(str_Full_Name, ".", "") 'Output the new user name, removes any periods after a middle initial

End Function
Public Function fx_Reverse_Name(strFullName)

' Purpose: This function reverses a users name that was LName, FName to be FName LName, (Ex. Rinaldi, Ethan -> Ethan Rinaldi)
' Trigger: Called
' Updated: 10/11/2019

' ***********************************************************************************************************************************

Dim strFirstName As String
    strFirstName = Right(strFullName, Len(strFullName) - InStrRev(strFullName, ",") - 1)

Dim strLastName As String
    strLastName = Left(strFullName, InStr(1, CStr(strFullName), ",") - 1)

fx_Reverse_Name = strFirstName & " " & strLastName

End Function
Public Function fx_Alphabetical_Sort_Collection(coll As Collection, intFirstRecordNum As Long, intLastRecordNum As Long) As Collection
  
' Purpose: This function will sort the values in an existing function alphabetically.
' Trigger: Called Function
' Updated: 5/3/2022
  
' Change Log:
'       5/3/2022:   Added the header info, and some comments
  
' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------
    
    Dim vTemp As Variant
  
    Dim intTempLow As Long
        intTempLow = intFirstRecordNum
    
    Dim intTempHigh As Long
        intTempHigh = intLastRecordNum
    
    Dim vCentreVal As Variant
        vCentreVal = coll((intFirstRecordNum + intLastRecordNum) \ 2)
  
' -----------------------
' Loop through the values
' -----------------------

    Do While intTempLow <= intTempHigh
  
        Do While coll(intTempLow) < vCentreVal And intTempLow < intLastRecordNum
            intTempLow = intTempLow + 1
        Loop
    
        Do While vCentreVal < coll(intTempHigh) And intTempHigh > intFirstRecordNum
            intTempHigh = intTempHigh - 1
        Loop
    
        If intTempLow <= intTempHigh Then
        
            'Swap values
            vTemp = coll(intTempLow)
          
            coll.Add coll(intTempHigh), After:=intTempLow
            coll.Remove intTempLow
          
            coll.Add vTemp, Before:=intTempHigh
            coll.Remove intTempHigh + 1
          
            'Move to next positions
            intTempLow = intTempLow + 1
            intTempHigh = intTempHigh - 1
          
        End If
    
    Loop
  
    ' Keep looping through the remaining values
    If intFirstRecordNum < intTempHigh Then fx_Alphabetical_Sort_Collection coll, intFirstRecordNum, intTempHigh
    If intTempLow < intLastRecordNum Then fx_Alphabetical_Sort_Collection coll, intTempLow, intLastRecordNum
    
    Set fx_Alphabetical_Sort_Collection = coll
  
End Function
Public Function fx_Find_String_In_Text_File(strFindMe As String, strFileLoc As String) As String
    
' Purpose: This function searches for a specific string in a text file, and outputs the location.
' Trigger: Called Function
' Updated: 9/16/2019

' ***********************************************************************************************************************************
    
' -----------
' Declare your variables
' -----------
    
    Dim objFSO As New FileSystemObject
    
    Dim txtFile As TextStream
        Set txtFile = objFSO.OpenTextFile(strFileLoc)
    
    Dim txtLine As String
    
    Dim strFindMeLoc As String
    
' -----------
' Find the text
' -----------
    
    Do Until txtFile.AtEndOfStream
        txtLine = txtFile.ReadLine
        If InStr(1, txtLine, strFindMe, vbTextCompare) > 0 Then
            strFindMeLoc = txtFile.Line
            Exit Do
        End If
    Loop

    txtFile.Close
    
    fx_Find_String_In_Text_File = strFindMeLoc
    
' -----------
' Release your variables
' -----------
    
    Set txtFile = Nothing
    Set objFSO = Nothing

End Function
Public Function fx_Get_First_Num(strSearch, intStart) As Long

' Purpose: This function finds the first number in a string, after the intStart, and returns the location.
' Trigger: Called Function
' Updated: 9/16/2019

' ***********************************************************************************************************************************

Dim i As Long
    
For i = intStart To Len(strSearch)
    If Mid(strSearch, i, 1) Like "#" Then
        fx_Get_First_Num = i
        Exit For
    End If
Next i

fx_Get_First_Num = i

End Function
Public Function fx_OpenWorkbook(wbPath As String) As Workbook
    
' Purpose: This function determines if a workbook is already open, otherwise open it.
' Trigger: Called Function
' Updated: 9/16/2019

' ***********************************************************************************************************************************
    
    Dim wb As Workbook
        
    Dim wbOpen As Workbook

    For Each wb In Workbooks
        If wb.FullName = wbPath Then Set wbOpen = wb
    Next wb
        
    If wbOpen Is Nothing Then Set wbOpen = Workbooks.Open(wbPath)
    
    Set fx_OpenWorkbook = wbOpen

End Function

Public Function fx_CloseWorkbook(wbPath As String)
    
' Purpose: This function closes the given workbook, based on the path.
' Trigger: Called Function
' Updated: 9/17/2019

' ***********************************************************************************************************************************
    
    Dim wb As Workbook
        
    Dim wbToClose As Workbook

    For Each wb In Workbooks
        If wb.FullName = wbPath Then Set wbToClose = wb
    Next wb
        
    If wbToClose Is Nothing Then Exit Function
    
    Application.DisplayAlerts = False
        wbToClose.Close
    Application.DisplayAlerts = True

End Function
Public Function fx_Create_Folder(strFullPath As String)

' Purpose: This function will create a folder if it doesn't already exist.
' Trigger: Called
' Updated: 9/18/2019

' ***********************************************************************************************************************************

    Dim objFSO As New FileSystemObject

    If Dir(strFullPath, vbDirectory) = vbNullString Then MkDir (strFullPath)
    
    fx_Create_Folder = objFSO.GetFolder(strFullPath)

End Function
Public Function fx_Pad_Zeros(strAccount, intAcctLength)

' Purpose: This function will pad the zeros in front of an account number, to the given length.
' Trigger: Called
' Updated: 9/19/2019

' ***********************************************************************************************************************************

Dim strTempAcct As String

If Len(strAccount) < intAcctLength Then
    strTempAcct = Application.WorksheetFunction.Rept("0", intAcctLength - Len(strAccount)) & strAccount
Else
    strTempAcct = strAccount
End If

fx_Pad_Zeros = strTempAcct

End Function
Public Function fx_Copy_to_Clipboard() As String

' Purpose: This function will copy the data into the applicable location.
' Trigger: Called Function
' Updated: 12/16/2020

' Change Log:
'       12/16/2020: Reduced the amount of code related to copying the selection, no longer used
'       12/24/2020: Added back in the copy paste from Seleciton in the Error Handler

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    Dim objDataObj As New MSForms.DataObject
    
    Dim strText As String
    
    Dim rng As Range
        Set rng = Selection
        
    Dim i As Long

' --------------------------------------
' Copy a selected range to the Clipboard
' --------------------------------------
    
    If rng.Rows.Count > 1 Then
    
        For i = rng.Row To rng.Row - 1 + rng.Rows.Count
            If Cells(i, rng.Column).Value <> "" Then
                strText = strText & Cells(i, rng.Column).Value & Chr(10) & Chr(10)
            End If
        Next i
        
        fx_Copy_to_Clipboard = Left$(strText, Len(strText) - 2) 'Remove trailing linebreaks
            Exit Function
    
    End If

' ---------------------
' Copy to the Clipboard
' ---------------------
    
    objDataObj.GetFromClipboard
        strText = Trim(objDataObj.GetText(1))
    
    'Application.CutCopyMode = False
    
    If Right$(strText, 2) = vbCrLf Or Right$(strText, 2) = vbNewLine Then strText = Left$(strText, Len(strText) - 2) 'Remove trailing linebreaks
    
    fx_Copy_to_Clipboard = strText

ErrorHandler:
   ' If strText = "" Then strText = Selection.Value2 '12/24/2021: Temporarily disabled
    'fx_Copy_to_Clipboard = strText

End Function
Public Function fx_Copy_from_Clipboard() As String

' Purpose: This function will copy the data from the clipboard and pass as a String.
' Trigger: Called Function
' Updated: 2/3/2022

' Change Log:
'       2/3/2022:   Reduced the amount of code related to copying the selection, no longer used

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim objDataObj As New MSForms.DataObject
        objDataObj.GetFromClipboard
    
    Dim strClipboardContents As String
        strClipboardContents = Trim(objDataObj.GetText(1))

' ---------------
' Pass the String
' ---------------
        
    fx_Copy_from_Clipboard = strClipboardContents
    
'    If Right$(strText, 2) = vbCrLf Or Right$(strText, 2) = vbNewLine Then strText = Left$(strText, Len(strText) - 2) 'Remove trailing linebreaks

End Function
Function fx_Privileged_User()

' Purpose: To output if the user is on the Privileged User list or not.
' Trigger: Called
' Updated: 2/15/2022

' Use Example: bolPrivilegedUser = fx_Privileged_User

' Change Log:
'       9/23/2020:  Intial Creation
'       12/17/2020: Added the conditional compiler constant to determine if DebugMode was on, if so make Priviledged User false.
'       1/30/2022:  Customized the list for the Deal Tracker
'       2/15/2022:  Passed the DebugMode Conditional Compilation Argument
    
' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strUserID As String
        strUserID = Application.UserName

    Dim bolPrivilegedUser As Boolean

' ------------------------------------
' Determine if the user is on the list
' ------------------------------------

    If _
        strUserID = "Rinaldi, James" Or _
        strUserID = "Rauckhorst, Eric W." Or _
        strUserID = "Duarte Espinoza, Axcel S." Or _
        strUserID = "Teti-Montanez, Robin" _
    Then
        bolPrivilegedUser = True
    Else
        bolPrivilegedUser = False
    End If

' ----------------
' Pass the results
' ----------------

#If DebugMode = 0 Then
    fx_Privileged_User = bolPrivilegedUser
#Else
    fx_Privileged_User = False
#End If

End Function

Public Sub fx_Incremental_Version_PPT()

' Purpose: This function backs up the applicable PPT to the (ARCHIVE) folder.
' Trigger: Quick Access Shortcut
' Updated: 1/28/2020

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    ' Declare your objects
    
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
    ' Declare your strings
    
    Dim pptName As String
        pptName = Left(ActivePresentation.Name, Len(ActivePresentation.Name) - 5)

    Dim strDir As String
        strDir = ActivePresentation.Path & "\"

    Dim strACFolder As String
        strACFolder = "K:\Audit\IA Leadership\Audit Committee\2020\2 - February 24, 2020\"
    
    Dim pptExt As String
        pptExt = Right(ActivePresentation.Name, 5)
        
    Dim pptFile As String
        pptFile = Dir(strDir & "*.ppt?")
    
    Dim strVersionCurrent As String
        strVersionCurrent = Mid(pptFile, InStrRev(pptFile, "v"), Len(pptFile))
   
    Dim strVersionNew As String
   
    ' Declare your integers
    
    Dim intVersionCurrent As Double
       intVersionCurrent = val(Replace(Replace(strVersionCurrent, pptExt, ""), "v", ""))

    Dim intVersionNew As Double

' --------------------------
' Find the latest itteration
' --------------------------
    
    Do While pptFile <> ""
        
        pptFile = Dir()
    
        If Not pptFile Like "*(MASTER)*" And pptFile <> "" Then
        
            strVersionNew = Mid(pptFile, InStrRev(pptFile, "v"), Len(pptFile))
            intVersionNew = val(Replace(Replace(strVersionNew, pptExt, ""), "v", ""))
            
            If intVersionNew > intVersionCurrent Then
                intVersionCurrent = intVersionNew
            End If

        End If

    Loop

    intVersionCurrent = intVersionCurrent + 0.1
        
' ----------------------------------
' Create a new iteration of the file
' ----------------------------------
    
    objFSO.CopyFile _
        Source:=ActivePresentation.FullName, _
        Destination:=strDir & "(DRAFT) February 2020 Audit Committee Deck v" & intVersionCurrent & pptExt
    
    objFSO.CopyFile _
        Source:=ActivePresentation.FullName, _
        Destination:=strACFolder & "(DRAFT) February 2020 Audit Committee Deck v" & intVersionCurrent & pptExt
    
End Sub


