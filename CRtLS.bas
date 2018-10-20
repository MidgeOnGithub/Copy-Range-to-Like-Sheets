Attribute VB_Name = "CRtLS"

'Copy Range to Like Sheets [CRtLS] Version 1.1

'AUTHOR:

    'Cody M Mason https://github.com/MidgeOnGithub

'LICENSE:

    'The MIT License (MIT)

    'Copyright (c) 2018 Cody M Mason

    'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

    'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

    'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

'DESCRIPTION:

    'CRtLS is an attempt at a user-friendly macro which guides the user through the process of copying values or formulas from a first worksheet into like-named sheets in a more immediate manner rather than manually copy-and-pasting into all sheets.

    'While several error-preventing methods have been implemented to prevent undesired outcomes or crashes from this macro, irregardless of the user's experience and knowledge with VBA and/or Excel, saving before running this macro is recommended, especially for large projects/worksheets. This is because its changes cannot be undone within a continuous instance of Excel. The user must re-load if something undesirable happens.

    'It can do so using either fixed formulas that set respective cells from other sheets equal to the cells in the first sheet, or by using the same relative formula from the first sheet, allowing for non-equal values but with samely calculations for all sheets.

    'In the case of referencing the value from the first sheet, running this macro just once forever permits the user to change *all sheets indicated to be edited by this macro* by only changing the first sheet.
        '-- If additional like-named sheets or values to be copied are added later, and the user wants them to also contain the value, the macro will need re-run. _
    'In the case of copying relative formulas, the macro must be re-run if the formula(s) on the first sheet is/are changed.

    'The user picks the "Prefix" by which the macro finds which sheets to edit.
    'The user may select their ActiveSheet as the "StartIndex," overriding the default first sheet with Prefix.
        '-- Useful if the first sheet(s) is/are different and don't want to be copied, perhaps because they are introductory or otherwise informative sheets.
    'The user may also elect a specific "EndIndex" by naming a sheet after which they'd like to stop the macro.
        '-- Useful if the final sheet(s) is/are different and don't want to be copied, perhaps because they contain summaries of the previous sheets.
    'The user chooses between copying either values or relative formulas.

    'There are several assumptions this macro makes: _
        'The user wants the *same range* from the first sheet to be edited on all other sheets.
            '-- The cell ranges do not adjust or offset as values/formula are copied to other worksheets. Therefore, the user needs to ensure the ranges are the same for all sheets before running this macro.
        'The user wants the *same value/formula* from the first sheet no all other sheets.
            '-- This macro will not increment, sum, or perform any similar modification to the values/formulas as it goes through successive worksheets.

    'This macro's usefulness extends to, but is not limited to, cases where the users wishes to edit numerous like-named sheets samely-format, such as inventory tabs, quantity sheets, or financial analysis sheets.

Option Explicit

Sub CRtLS()
    Dim TWB As Workbook: Set TWB = ThisWorkbook
    Dim ASht As Worksheet: Set ASht = TWB.ActiveSheet
    Dim intWS As Long: intWS = TWB.Worksheets.Count
    Dim strASht As String: strASht = ASht.Name
    Dim intASht As Long: intASht = ASht.Index
    'Shorthands declared and set! =================================================
    
    'Welcome and guidance message.
    MsgBox "This macro will only operate on sheets with specified names. It assumes you want to start from the sheet where you called the macro.", vbInformation, "Welcome to the CRtLS Macro!"

    Dim SearchType As Variant

    Do
        SearchType = InputBox("How should we search for sheet names?" & vbCr & vbCr & "0 = All Sheets" & vbCr & "1 = Prefix Only" & vbCr & "2 = Suffix Only" & vbCr & "3 = Both Prefix and Suffix", "Prefix, Suffix, or Both?", "1")
    Loop Until SearchType >= 0 And SearchType <= 3

    If MsgBox("0 = All Sheets" & vbCr & "1 = Prefix Only" & vbCr & "2 = Suffix Only" & vbCr & "3 = Both Prefix and Suffix" & vbCr & vbCr & "Response interpreted as " & SearchType & ". Is this ok?", vbQuestion + vbOKCancel, "Confirm Input") = vbCancel Then GoTo Cancelled
    'Determined if searching by prefix and/or suffix! =============================
    
    Dim MsgFlag As Integer 'For Cancelled section messages.

    If SearchType = 0 Then GoTo WorkZone 'No name searching needed for SearchType = 0.
    
    Dim Prefix As String

    If SearchType = 1 Or SearchType = 3 Then 'Values corresponding to Prefix Only and Both will run this.
        'User Input of Prefix
        Prefix = InputBox(Prompt:="Please enter the letter/number/symbol that prefixes sheets you wish to update. Case and Space Sensitive.", Title:="Input Desired Sheet Prefix")

        If Prefix = "" Then 'If user enters nothing or escapes out of Inputbox with no input given.
            GoTo Cancelled
        ElseIf InStr(strASht, Prefix) <> 1 Then
            MsgFlag = 10
            GoTo Cancelled
        End If
    End If

    If SearchType >= 2 Then 'Values of 2 or 3 will run this.
        Dim Suffix As String 'User Input of Suffix
        Suffix = InputBox(Prompt:="Please enter the letter/number/symbol that suffixes sheets you wish to update. Case and Space Sensitive.", Title:="Input Desired Sheet Suffix")
    
        Dim SuffixFromEnd As Integer 'Needed when doing worksheet name checks.
        SuffixFromEnd = Len(Suffix) - 1 'Tells what position to use when comparing names of sheets to Suffix. If suffix length is 1, end position should equal length of sheet name, therefore, subtract 1.
    
        If Suffix = "" Then 'If user enters nothing or escapes out of Inputbox with no input given.
            GoTo Cancelled
        ElseIf InStr(strASht, Suffix) <> Len(strASht) - SuffixFromEnd Then
            MsgFlag = 10
            GoTo Cancelled
        End If
    End If
    'Prefix and/or Suffix Set! ====================================================

    Dim o As Long, StartIndex As Long

    If SearchType = 1 Or SearchType = 3 Then 'Values corresponding to Prefix Only and Both will run this.
        For o = 1 To intWS
            If InStr(TWB.Worksheets(o).Name, Prefix) = 1 Then
                StartIndex = o 'Found the first sheet with Prefix at beginning.
                Exit For 'Go back to Sub
            ElseIf o >= intWS And InStr(TWB.Worksheets(o).Name, Prefix) <> 1 Then 'In case the For Each tries all sheets but none have Prefix.
                MsgFlag = 30
                GoTo Cancelled
            End If
        Next o

        If StartIndex <> intASht Then 'Checks if first sheet with Prefix doesn't equal current sheet.
            'Because the sub uses the Active Sheet to set MatchRge, this If attempts to prevents the user from changing something from a "wrong" sheet.
            If MsgBox("You are not currently in the first sheet with the prefix " & Prefix & "." & vbCr & vbCr & "If you intended to run the macro from sheet " & strASht & " on, hit OK.", vbInformation + vbOKCancel, "Start Sheet Confirmation") = vbCancel Then
                GoTo Cancelled
            Else
                StartIndex = intASht
            End If
        End If 'If the variables are equal, then no need to sniff for errors.
    End If 'End of Prefix If.

    If SearchType >= 2 Then 'Values corresponding to Suffix Only or Both will run this.
        For o = 1 To intWS
            If InStr(TWB.Worksheets(o).Name, Suffix) = Len(TWB.Worksheets(o).Name) - SuffixFromEnd Then
                StartIndex = o 'Found the first sheet with Suffix at end.
                Exit For 'Go back to Sub
            ElseIf o >= intWS And InStr(TWB.Worksheets(o).Name, Suffix) <> 1 Then 'In the case the For Each goes through all sheets but none have the suffix.
                MsgFlag = 40
                GoTo Cancelled
            End If
        Next o
    
        If StartIndex <> intASht Then 'Checks if first sheet with Suffix doesn't equal current sheet.
            'Because the sub uses the Active Sheet for MatchRge, this If attempts to prevents the user from changing something from a "wrong" sheet.
            If MsgBox("You are not currently in the first sheet with the suffix " & Suffix & "." & vbCr & vbCr & "If you intended to run the macro from sheet " & strASht & " on, hit OK.", vbInformation + vbOKCancel, "Start Sheet Confirmation") = vbCancel Then
                GoTo Cancelled
            Else
                StartIndex = intASht
            End If
        End If 'If the variables are equal, then no need to sniff for errors.
    End If 'End of Suffix If

    Dim StartWS As Worksheet: Set StartWS = TWB.Worksheets(StartIndex)
    'StartWS successfully validated! ==============================================

    WorkZone: 'Section indicator allows SearchType = 0 to bypass unnecessary code.
    'In case users have sheets with Prefix/Suffix towards the end of their workbook which are different and/or they don't wish to change. One of the limitations of this macro is that it can only be given one start and end point.
    Dim EndIndex As Long
    Dim EndWS As Worksheet
    Dim SearchString As String 'Used in the WhereEnd function if user wants the macro to halt after a certain sheet.

    Select Case SearchType 'To determine EndIndex value.
        Case 0 'Doesn't need to call the function, just assign EndIndex to last sheet.
            If intASht <> 1 Then 'Checks if user is actually on the first sheet since they indicated to change all sheets.
                If MsgBox("You indicated All Sheets, but are not calling this macro from the workbook's first sheet. This means only sheets from " & strASht & " on will be edited using information from said sheet." & vbCr & vbCr & "If this was your intention, hit OK to proceed.", vbInformation + vbOKCancel, "Confirm Input") = vbCancel Then GoTo Cancelled
            End If
        
            StartIndex = intASht
            Set StartWS = TWB.Worksheets(StartIndex)
            EndIndex = intWS
        Case 1
            EndIndex = FindEnd(TWB, intWS, Prefix, "P", StartIndex) 'P indicates we want FindEnd to search sheet names considering the SearchString as a prefix.
            SearchString = Prefix
        Case 2
            EndIndex = FindEnd(TWB, intWS, Suffix, "S", StartIndex) 'S indicates we want FindEnd to search sheet names considering the SearchString as a suffix
        Case 3
            Dim End1 As Integer
            Dim End2 As Integer 'Now we are doing two searches, and we want to compare the results in case of differences.
        
            End1 = FindEnd(TWB, intWS, Prefix, "P", StartIndex)
            End2 = FindEnd(TWB, intWS, Suffix, "S", StartIndex)
        
            If End1 >= End2 Then 'Assign EndIndex to the larger of the End1 and End2.
                EndIndex = End1
            Else
                EndIndex = End2
            End If
    End Select

    Set EndWS = TWB.Worksheets(EndIndex)

    If EndIndex <= StartIndex Then 'Prevents errors by stopping if StartIndex is the last applicable sheet. Macro is useless in such a case.
        MsgFlag = 20
        GoTo Cancelled
    ElseIf MsgBox("The CRtLS Macro will work from sheets " & StartWS.Name & " to " & EndWS.Name & "." & vbCr & vbCr & "If this is ok, press Yes." & vbCr & "If you have a specific sheet after which the macro should stop, press No.", vbYesNo, "Ending Sheet Question") = vbNo Then
        EndIndex = WhereEnd(TWB, intWS, strASht, StartIndex) 'They want to stop somewhere in particular: WhereEnd finds where.
        If EndIndex = -1 Then GoTo Cancelled 'If the users selects Cancel whilst WhereEnd is running, WhereEnd returns -1
    End If
    'EndIndex successfully determined. ============================================

    Dim MatchRge As Range

    Do
        On Error GoTo Cancelled
        'Application.InputBox allows for standard Excel UX for range selection. Type 8 indicates input must be a Range.
        Set MatchRge = Application.InputBox(Prompt:="Use your mouse and/or keyboard to enter a range from sheet " & strASht & ".", Title:="Input Range of Cells", Type:=8)

        If MatchRge.Worksheet.Index <> intASht Then
            MsgBox "Selected range must be within sheet " & strASht & " in workbook " & TWB.Name & "."
        End If
    Loop Until MatchRge.Worksheet.Index = intASht

    'Shorthands
    Dim RgeRows As Integer: RgeRows = MatchRge.Rows.Count
    Dim RgeCols As Integer: RgeCols = MatchRge.Columns.Count

    'Creates an array and then sizes it to dimensions of MatchRge
    Dim WhatToCopy() As Variant 'Must offset because arrays are indexed from 0 while Ranges are from 1.
    ReDim Preserve WhatToCopy(RgeRows - 1, RgeCols - 1)
    'Successfully got MatchRge and sized WhatToCopy. ==============================

    Dim ExtractType As Integer 'Determine whether to copy values or formulas

    Do
        ExtractType = Application.InputBox(Prompt:="What should we do to range " & MatchRge.Address & " in other worksheets?" & vbCr & vbCr & "1 = Copy Values Only" & vbCr & "2 = Set Cells Equal to " & StartWS.Name & " Cells" & vbCr & "3 = Copy Relative Formulas", Title:="What Should We Do?", Default:="1", Type:=2)
    Loop Until ExtractType >= 1 And ExtractType <= 3

    If MsgBox("1 = Copy Values Only" & vbCr & "2 = Set Cells Equal to " & StartWS.Name & " Cells" & vbCr & "3 = Copy Relative Formulas" & vbCr & vbCr & "Response interpreted as " & ExtractType & ". Is this ok?", vbQuestion + vbOKCancel, "Confirm Input") = vbCancel Then GoTo Cancelled
    'Successfully determined what to do with data from MatchRge. ==================

    'Iterate through MatchRge to plug values/formulas into WhatToCopy array.
    Dim i As Integer, j As Integer

    For i = 1 To RgeRows
        For j = 1 To RgeCols
            'Must offset WhatToCopy index by 1 because array is indexed from 0 while Range is from 1.
            Select Case ExtractType
                Case 1
                    WhatToCopy(i - 1, j - 1) = MatchRge(i, j).Value
                Case 2
                    WhatToCopy(i - 1, j - 1) = MatchRge(i, j).Address
                    'Now add StartWS reference, ultimately turning making formulas setting cells equal to the MatchRge.
                    WhatToCopy(i - 1, j - 1) = "='" & StartWS.Name & "'!" & WhatToCopy(i - 1, j - 1)
                Case 3
                    If MatchRge(i, j).Formula = "" Then
                        WhatToCopy(i - 1, j - 1) = ""
                    Else 'Prevents just writing an equals sign into another sheet's cell.
                        WhatToCopy(i - 1, j - 1) = "=" & MatchRge(i, j).Formula
                    End If
            End Select
        Next j
    Next i

    'Successfully populated WhatToCopy. ===========================================

    Dim k As Integer, kWS As Worksheet
    Dim arrShtsSkip() As Variant
    Dim intSkip As Long: intSkip = 0

    'The loop that actually does the editing.
    For k = StartIndex + 1 To EndIndex
        Set kWS = TWB.Worksheets(k)
    
        Select Case SearchType
            Case 1 'Will only apply changes to sheets with Prefix in the beginning of their name.
                If Not InStr(kWS.Name, Prefix) = 1 Then
                    ReDim Preserve arrShtsSkip(UBound(arrShtsSkip) + 1) 'Increase array size by one.
                    arrShtsSkip(UBound(arrShtsSkip)) = kWS.Name 'Place WS name in end of array.
                    GoTo NextSht
                End If
            Case 2
                If Not InStr(kWS.Name, Suffix) = Len(Suffix) - 1 Then
                    ReDim Preserve arrShtsSkip(UBound(arrShtsSkip) + 1)
                    arrShtsSkip(UBound(arrShtsSkip)) = kWS.Name
                    GoTo NextSht
                End If
            Case 3
                If Not InStr(kWS.Name, Prefix) = 1 Or InStr(kWS.Name, Suffix) = Len(Suffix) - 1 Then
                    ReDim Preserve arrShtsSkip(intSkip)
                    arrShtsSkip(intSkip) = kWS.Name
                    intSkip = intSkip + 1
                    GoTo NextSht
                End If
        End Select
    
        For i = 1 To RgeRows
            For j = 1 To RgeCols
                kWS.Range(MatchRge(i, j).Address) = WhatToCopy(i - 1, j - 1)
            Next j
        Next i
    NextSht:
    Next k
    'Successfully input data into indicated sheets! ===============================

    Dim ArrDimensions As Integer
    ArrDimensions = ArrayDimensionCount(arrShtsSkip)

    'Will show user the sheets skipped during operation, if desired.
    If ArrDimensions > 0 Then 'If array is dimensionless (ArrDimensions = 0), no sheets were skipped.
        If MsgBox("Would you like to see the list of sheets skipped during processing?", vbQuestion + vbYesNo, "See Sheets Skipped?") = vbYes Then
            Dim ShtsSkipList As String: ShtsSkipList = Join(arrShtsSkip, ", ")
            MsgBox "Sheets skipped:" & vbCr & vbCr & ShtsSkipList, vbInformation + vbOKOnly, "List of Sheets Skipped"
        End If
    End If

    Dim strExtractTypeMsg As String

    Select Case ExtractType
        Case 1
            strExtractTypeMsg = "values matching "
        Case 2
            strExtractTypeMsg = "corresponding ranges set to equal cells from "
        Case 3
            strExtractTypeMsg = "relative formulas matching "
    End Select

    MsgBox "Indicated sheets from " & StartWS.Name & " to " & EndWS.Name & " now have " & strExtractTypeMsg & StartWS.Name & " for range " & MatchRge.Address & ".", vbExclamation + vbOKOnly, "Success!"

    End
    'Complete! Unless somehow the macro is cancelled... ===========================

    Cancelled:
    Dim MsgTxt As String
    Dim vbType As VbMsgBoxStyle

    Select Case MsgFlag
        Case 10
            MsgTxt = "You need to call this macro from a sheet with the indicated prefix/suffix in its name." & vbCr & vbCr & ""
            vbType = vbInformation
        Case 20
            MsgTxt = "Current sheet is the last sheet with like name, macro is unnecessary." & vbCr & vbCr & ""
            vbType = vbInformation
        Case 30
            MsgTxt = "Failed to find any sheet starting with " & Prefix & "." & vbCr & vbCr & ""
            vbType = vbCritical
        Case 40
            MsgTxt = "Failed to find any sheet ending with " & Suffix & "." & vbCr & vbCr & ""
            vbType = vbCritical
        Case Else
            MsgTxt = ""
            vbType = vbInformation
    End Select

    MsgBox MsgTxt & "Macro cancelled. No edits made.", vbType + vbOKOnly, "Macro Cancelled"

End Sub

Private Function ArrayDimensionCount(ArrayOfInterest As Variant) As Integer
    Dim NumOfDims As Integer: NumOfDims = 0
    Dim LastIndex As Integer

    On Error Resume Next 'Will exit loop upon "LastIndex =..." line throwing an error.
    Do Until Not Err.Number = 0
        NumOfDims = NumOfDims + 1
        LastIndex = UBound(ArrayOfInterest, NumOfDims) 'This line throws an error if the dimension doesn't exist.
    Loop

    ArrayDimensionCount = NumOfDims - 1 'Go back to the last dimension count that existed.

End Function

Private Function FindEnd(TWB As Workbook, intWS As Long, SearchString As String, SearchType As String, StartIndex As Long) As Long
    Dim i As Long

    Select Case SearchType
        Case "P"
            For i = StartIndex To intWS
                If InStr(TWB.Worksheets(i).Name, SearchString) = 1 Then FindEnd = i
            Next i
        Case "S"
            For i = StartIndex To intWS
                If InStr(TWB.Worksheets(i).Name, SearchString) = Len(SearchString) - 1 Then FindEnd = i
            Next i
    End Select

End Function

Private Function WhereEnd(TWB As Workbook, intWS As Long, strASht As String, StartIndex As Long) As Long
    Dim i As Long
    Dim Confirm As Variant
    Dim UserIndication As Variant

    Do
        UserIndication = InputBox(Prompt:="Type in exactly the name of the sheet after which you want the macro to stop. Case and Space Sensitive.", Title:="Indicate End Sheet's Name", Default:=strASht)
        
        If StrPtr(UserIndication) = 0 Then 'Cancel if they cancel.
            WhereEnd = -1
            Exit Function
        End If
    
        For i = 1 To intWS 'Loop to find index of sheet.
            If TWB.Worksheets(i).Name = UserIndication Then
                If i <= StartIndex Then 'Checks if indicated sheet is not after start sheet; prompt user again if so.
                    MsgBox "Found " & UserIndication & ", but it doesn't come after " & strASht & "", vbOKOnly, "Invalid Sheet"
                    Exit For
                Else
                    Confirm = MsgBox("Found sheet " & UserIndication & " with index " & i & " out of " & intWS & " total sheets. Confirm with Yes or select No to search another sheet name.", vbInformation + vbYesNo, "Confirm Sheet to End")

                    If Confirm = vbYes Or Confirm = vbNo Then
                        Exit For
                    ElseIf Confirm = vbCancel Then 'Cancel if they cancel.'
                        WhereEnd = -1
                        Exit Function
                    End If
                End If

            ElseIf i = intWS Then 'Alerts user if all sheets were searched and no name match was found.
                MsgBox "Sheet name not found. This is case sensitive and requires an exact match.", vbCritical
                Exit For
            End If
        Next i
    Loop Until TWB.Worksheets(i).Name = UserIndication And Confirm = vbYes

    WhereEnd = i

End Function
