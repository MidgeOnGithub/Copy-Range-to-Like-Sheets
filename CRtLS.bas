Attribute VB_Name = "CRtLS"
'Copy Range to Like Sheets [CRtLS] Version 1.0


'AUTHOR: Cody M Mason https://github.com/MidgeOnGithub


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
'   -- If additional like-named sheets or values to be copied are added later, and the user wants them to also contain the value, the macro will need re-run. _
'In the case of copying relative formulas, the macro must be re-run if the formula(s) on the first sheet is/are changed.

'The user picks the "Prefix" by which the macro finds which sheets to edit.
'The user may select their ActiveSheet as the "StartIndex," overriding the default first sheet with Prefix.
'   -- Useful if the first sheet(s) is/are different and don't want to be copied, perhaps because they are introductory or otherwise informative sheets.
'The user may also elect a specific "EndIndex" by naming a sheet after which they'd like to stop the macro.
'   -- Useful if the final sheet(s) is/are different and don't want to be copied, perhaps because they contain summaries of the previous sheets.
'The user chooses between copying either values or relative formulas.

'There are several assumptions this macro makes: _
    'The user wants the *same range* from the first sheet to be edited on all other sheets.
    '   -- The cell ranges do not adjust or offset as values/formula are copied to other worksheets. Therefore, the user needs to ensure the ranges are the same for all sheets before running this macro.
    'The user wants the *same value/formula* from the first sheet no all other sheets.
    '   -- This macro will not increment, sum, or perform any similar modification to the values/formulas as it goes through successive worksheets.

'This macro's usefulness extends to, but is not limited to, cases where the users wishes to edit numerous like-named sheets samely-format, such as inventory tabs, quantity sheets, or financial analysis sheets.

Option Explicit
Sub CRtLS()

'Welcome and guidance message.
MsgBox "This macro will only operate on sheets with specified names. You will be given a series of prompts to guide you through the macro, with multiple opportunities to cancel if something looks wrong." & vbCr & vbCr & "The following prompt determines if you wish to identify sheets by prefix, suffix, or both.", vbInformation, "Welcome to the CRtLS Macro!"

Dim SearchType As Variant

Do
    SearchType = InputBox("How should we search for sheet names?" & vbCr & vbCr & "0 = All Sheets" & vbCr & "1 = Prefix Only" & vbCr & "2 = Suffix Only" & vbCr & "3 = Both Prefix and Suffix", "Prefix, Suffix, or Both?", "1")

Loop Until SearchType >= 0 And SearchType <= 3

If MsgBox("0 = All Sheets" & vbCr & "1 = Prefix Only" & vbCr & "2 = Suffix Only" & vbCr & "3 = Both Prefix and Suffix" & vbCr & vbCr & "Response interpreted as " & SearchType & ". Is this ok?", vbQuestion + vbOKCancel, "Confirm Input") = vbCancel Then
    GoTo Cancelled

End If

'Determined if searching by prefix and/or suffix! =============================

Dim TWB As Workbook, intWS As Integer, ASht As Worksheet, strASht As String, intASht As Integer 'Shorthands.

Set TWB = ThisWorkbook
intWS = TWB.Worksheets.Count
Set ASht = TWB.ActiveSheet
strASht = ASht.Name
intASht = ASht.Index

'Shorthands declared and set! =================================================

Dim MsgFlag As Integer 'For Cancelled section messages.

If SearchType = 0 Then GoTo WorkZone 'No name searching needed for SearchType = 0.
    
Dim Prefix As String

If SearchType = 1 Or SearchType = 3 Then 'Values corresponding to Prefix Only and Both will run this.

    'User Input of Prefix
    Prefix = InputBox(Prompt:="Please enter the letter/number/symbol that prefixes sheets you wish to update. Case and Space Sensitive.", Title:="Input Desired Sheet Prefix")

    If StrPtr(Prefix) = 0 Then 'StrPtr = 0 if user cancels or otherwise escapes out of Inputbox with no input given.
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
    
    If StrPtr(Suffix) = 0 Then 'StrPtr = 0 if user cancels or otherwise escapes out of Inputbox.
        GoTo Cancelled

    ElseIf InStr(strASht, Suffix) <> Len(strASht) - SuffixFromEnd Then
        MsgFlag = 10
        GoTo Cancelled
    
    End If
    
End If

'Prefix and/or Suffix Set! ====================================================

Dim o As Integer
    
If SearchType = 1 Or SearchType = 3 Then 'Values corresponding to Prefix Only and Both will run this.

    Dim intFirstPSht As Integer 'Indicates first Prefix sheet index.
    
    For o = 1 To intWS
    
        If InStr(TWB.Worksheets(o).Name, Prefix) = 1 Then
            intFirstPSht = o 'Found the first sheet with Prefix at beginning.
            Exit For 'Go back to Sub

        ElseIf o >= intWS And InStr(TWB.Worksheets(o).Name, Prefix) <> 1 Then 'In case the For Each tries all sheets but none have Prefix.
            MsgFlag = 30
            GoTo Cancelled

        End If
    
    Next o

    If intFirstPSht <> intASht Then 'Checks if first sheet with prefix doesn't equal current sheet's prefix.

        'Because the sub uses the Active Sheet to set MatchRge, this If attempts to prevents the user from changing something from a "wrong" sheet.

        If MsgBox("You are not currently in the first sheet with the prefix " & Prefix & "." & vbCr & vbCr & "If you intended to run the macro from sheet " & strASht & " on, hit OK." & vbCr & vbCr & "You may Cancel and re-call this macro to run from the first applicable sheet, if desired.", vbOKCancel, "Start Sheet Discrepancy") = vbCancel Then
            GoTo Cancelled
       
        Else
            intFirstPSht = intASht

        End If
    
    Dim strFirstPSht As String 'Indicates first Prefix sheet name.
    strFirstPSht = TWB.Worksheets(intFirstPSht).Name
    
    End If 'If the variables are equal, then no need to sniff for errors.

End If 'End of Prefix If.

If SearchType >= 2 Then 'Values corresponding to Suffix Only or Both will run this.

    Dim intFirstSSht As Integer 'Indicates first Suffix sheet index.

    For o = 1 To intWS
    
        If InStr(TWB.Worksheets(o).Name, Suffix) = Len(TWB.Worksheets(o).Name) - SuffixFromEnd Then
            intFirstSSht = o 'Found the first sheet with Suffix at end.
            Exit For 'Go back to Sub

        ElseIf o >= intWS And InStr(TWB.Worksheets(o).Name, Suffix) <> 1 Then 'In the case the For Each goes through all sheets but none have the suffix.
            MsgFlag = 40
            GoTo Cancelled
        
        End If
    
    Next o
    
    '!!!Code to check ASht first vs FirstSSht goes here.
    
    Dim strFirstSSht As String 'Indicates first Suffix sheet name.
    strFirstSSht = TWB.Worksheets(intFirstSSht).Name
    
End If 'End of Suffix If

Dim StartIndex As Integer

If intFirstPSht = intFirstSSht Then
    StartIndex = intFirstPSht
    
Else
    '!!!Code to decided between intFirstPSht & intFirstSSht goes here.
        
End If

Dim StartSht As Worksheet
Set StartSht = TWB.Worksheets(StartIndex)

'StartIndex successfully determined! ==========================================

WorkZone: 'Section indicator to allow SearchType = 0 to skip through unnecessary code.

'In case users have sheets with Prefix/Suffix towards the end of their workbook which are different and/or they don't wish to change. One of the limitations of this macro is that it can only be given one start and end point.
Dim EndIndex As Integer
Dim EndWS As Worksheet
Dim SearchString As String 'Used in the WhereEnd function if user wants the macro to halt after a certain sheet.

Select Case SearchType 'To determine EndIndex value.

    Case 0 'Doesn't need to call the function, just assign EndIndex to last sheet.
        
        If intASht <> 1 Then 'Checks if user is actually on the first sheet since they indicated to change all sheets.
        
            If MsgBox("You indicated All Sheets, but are not calling this macro from the workbook's first sheet. This means only sheets from " & strASht & " on will be edited using information from said sheet." & vbCr & vbCr & "If this was your intention, hit OK to proceed.", vbInformation + vbOKCancel, "Confirm Input") = vbCancel Then
                GoTo Cancelled

            End If
    
        End If
        
        StartIndex = intASht
        EndIndex = intWS
        
    Case 1
        EndIndex = FindEnd(TWB, intWS, Prefix, "P", StartIndex) 'P indicates we want FindEnd to search sheet names considering the SearchString as a prefix.
        SearchString = Prefix
        
    Case 2
        EndIndex = FindEnd(TWB, intWS, Suffix, "S", StartIndex) 'S indicates we want FindEnd to search sheet names considering the SearchString as a suffix.
    
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

If EndIndex <= StartIndex Then 'Prevents errors by not calling further processes if StartIndex is the last sheet. Macro is useless in such a case.
    MsgFlag = 20
    GoTo Cancelled
    
ElseIf MsgBox("The CRtLS Macro will work from sheets " & StartWS.Name & " to " & EndWS.Name & "." & vbCr & vbCr & "If this is ok, press Yes." & vbCr & "If you have a specific sheet after which the macro should stop, press No.", vbYesNo, "Ending Sheet Question") = vbNo Then
    EndIndex = WhereEnd(TWB, intWS, strASht, EndIndex) 'They want to stop somewhere in particular: WhereEnd finds where.
    
    If EndIndex = -1 Then 'If the users selects Cancel whilst WhereEnd is running, WhereEnd returns -1.
        GoTo Cancelled
        
    End If
    
End If
      
'EndIndex successfully determined. ============================================

Dim MatchRge As Range

'Application.InputBox allows for standard Excel UX for range selection. Type 8 indicates input must be a Range.
Set MatchRge = Application.InputBox(Prompt:="Use your mouse and/or keyboard to enter a range of cells from sheet " & strASht & ".", _
                                    Title:="Input Range of Cells", Type:=8)

'!!!Code to check if selected range is on ASht goes here.

'Shorthands
Dim RgeRows As Integer
Dim RgeCols As Integer

RgeRows = MatchRge.Rows.Count
RgeCols = MatchRge.Columns.Count

'Creates an array and then sizes it to dimensions of MatchRge
Dim WhatToCopy() As Variant
ReDim WhatToCopy(RgeRows - 1, RgeCols - 1) 'Must subtract because arrays are indexed from 0 while Ranges are from 1.

'Successfully got MatchRge and sized WhatToCopy. ==============================

Dim ExtractType As Variant 'Determine whether to copy values or formulas

Do
    ExtractType = InputBox("What should we do to other worksheet's cells corresponding to selected range " & MatchRge.Address & "?" & vbCr & vbCr & "1 = Copy Values Only" & vbCr & "2 = Set Cells Equal to " & StartWS.Name & " Cells" & vbCr & "3 = Copy Relative Formulas", "What Should We Do?", "1")

Loop Until ExtractType >= 1 And ExtractType <= 3

If MsgBox("1 = Copy Values Only" & vbCr & "2 = Set Cells Equal to " & StartWS.Name & " Cells" & vbCr & "3 = Copy Relative Formulas" & vbCr & vbCr & "Response interpreted as " & ExtractType & ". Is this ok?", vbQuestion + vbOKCancel, "Confirm Input") = vbCancel Then
    GoTo Cancelled

End If

'Successfully determined if extracting values or formulas. =======================================

'Iterate through MatchRge to plug values/formulas into WhatToCopy array.

Dim i As Integer, j As Integer

For i = 1 To RgeRows
    
    For j = 1 To RgeCols
        
        'Must offset WhatToCopy value by 1 because array is indexed from 0 while range is from 1.
        Select Case ExtractType
        
            Case ExtractType = 1
                WhatToCopy(i - 1, j - 1) = MatchRge(i, j).Value
            
            Case ExtractType = 2
                WhatToCopy(i - 1, j - 1) = MatchRge(i, j).Address
                
                'Now make the values of WhatToCopy into formula settings cells equal by adding StartIndex sheet.
                WhatToCopy(i - 1, j - 1) = "='" & StartWS.Name & "'!" & WhatToCopy(i - 1, j - 1)
            
            Case ExtractType = 3
                WhatToCopy(i - 1, j - 1) = MatchRge(i, j).Formula
            
                If WhatToCopy(i - 1, j - 1) <> "" Then 'Prevents just writing an equals sign into another sheet's cell, just leaves it blank like the parent cell.
                    WhatToCopy(i - 1, j - 1) = "=" & WhatToCopy(i - 1, j - 1)
                
                End If
        
        End Select
        
    Next j
    
Next i

'Successfully populated WhatToCopy. ===========================================

Dim k As Integer

'The loop that actually does the editing.
For k = StartIndex + 1 To EndIndex

    If InStr(TWB.Worksheets(k).Name, Prefix) <> 1 Then
        GoTo NextIteration
        
    End If
    
    For i = 1 To RgeRows
    
        For j = 1 To RgeCols
        
            TWB.Worksheets(k).Range(MatchRge(i, j).Address) = WhatToCopy(i - 1, j - 1)
        
        Next j
        
    Next i
    
NextIteration:
    
Next k

'Successfully input data into indicated sheets. ===============================

Dim strExtractTypeMsg As String

Select Case ExtractType
    
    Case ExtractType = 1
        strExtractTypeMsg = "values matching "
    
    Case ExtractType = 2
        strExtractTypeMsg = "corresponding ranges set to equal cells from "
    
    Case ExtractType = 3
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

Private Function FindEnd(TWB As Workbook, intWS As Integer, SearchString As String, SearchType As String, StartIndex As Integer) As Integer

Dim i As Long

For i = StartIndex To intWS

    If InStr(TWB.Worksheets(i).Name, SearchString) = 1 Then
        FindEnd = i
        
    End If
    
Next i

End Function

Private Function WhereEnd(TWB As Workbook, intWS As Integer, strASht As String, StartIndex As Integer) As Integer

Dim i As Long
Dim Confirm As Variant
Dim UserIndication As Variant

Do
    UserIndication = InputBox(Prompt:="Type in exactly the name of the sheet after which you want the macro to stop. Case and Space Sensitive.", _
                              Title:="Indicate End Sheet's Name")

    If StrPtr(UserIndication) = 0 Then GoTo Cancelled 'Cancel if they cancel.
    
    For i = 1 To intWS 'Inner loop begins to find index of sheet.
    
        If TWB.Worksheets(i).Name = UserIndication Then
            
            If i <= StartIndex Then 'Checks if indicated sheet is not after start sheet; prompt user again if so.
                MsgBox "Worksheet found, but either equals or comes before " & strASht & "", vbOKOnly, "Invalid Sheet"
                Exit For
            
            Else
                Confirm = MsgBox("Found sheet " & UserIndication & " with index " & i & " out of " & intWS & " total sheets. Confirm with Yes or select No to search another sheet name.", vbInformation + vbYesNo, "Confirm Sheet to End")
            
                If Confirm = vbYes Or Confirm = vbNo Then
                    Exit For
            
                ElseIf Confirm = vbCancelled Then 'Cancel if they cancel.'
                    GoTo Cancelled
                       
            End If
            
        ElseIf i = intWS Then 'Alerts user if all sheets were searched and no name match was found.
            MsgBox "Sheet name not found. This is case sensitive and requires an exact match.", vbCritical
            Exit For
               
        End If
        
    Next i 'Inner loop ends.

Loop Until TWB.Worksheets(i).Name = UserIndication And Confirm = vbYes

WhereEnd = i

Exit Function

Cancelled:
    WhereEnd = -1

End Function
