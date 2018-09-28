Attribute VB_Name = "CRtLS"
'Copy Range to Like Sheets [CRtLS] Version 1.0


'AUTHOR: Cody M Mason https://github.com/MidgeOnGithub


'LICENSE: _
 _
The MIT License (MIT) _
Copyright (c) 2018 Cody M Mason _
 _
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: _
 _
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. _
 _
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


'DESCRIPTION: _
CRtLS is an attempt at a user-friendly macro which guides the user through the process of copying values or formulas from a first worksheet into like-named sheets in a more immediate manner rather than manually copy-and-pasting into all sheets. _
 _
While several error-preventing methods have been implemented to prevent undesired outcomes or crashes from this macro, irregardless of the user's experience and knowledge with VBA and/or Excel, saving before running this macro is recommended, especially for large projects/worksheets. This is because its changes cannot be undone within a continuous instance of Excel. The user must re-load if something undesirable happens. _
 _
It can do so using either fixed formulas that set respective cells from other sheets equal to the cells in the first sheet, or by using the same relative formula from the first sheet, allowing for non-equal values but with samely calculations for all sheets. _
 _
In the case of referencing the value from the first sheet, running this macro just once forever permits the user to change *all sheets indicated to be edited by this macro* by only changing the first sheet. _
    -- If additional like-named sheets or values to be copied are added later, and the user wants them to also contain the value, the macro will need re-run. _
In the case of copying relative formulas, the macro must be re-run if the formula(s) on the first sheet is/are changed. _
 _
The user picks the "Prefix" by which the macro finds which sheets to edit. _
The user may select their ActiveSheet as the "StartIndex," overriding the default first sheet with Prefix. _
    -- Useful if the first sheet(s) is/are different and don't want to be copied, perhaps because they are introductory or otherwise informative sheets. _
The user may also elect a specific "EndIndex" by naming a sheet after which they'd like to stop the macro. _
    -- Useful if the final sheet(s) is/are different and don't want to be copied, perhaps because they contain summaries of the previous sheets. _
The user chooses between copying either values or relative formulas. _
 _
There are several assumptions this macro makes: _
    The user wants the *same range* from the first sheet to be edited on all other sheets. _
        -- The cell ranges do not adjust or offset as values/formula are copied to other worksheets. Therefore, the user needs to ensure the ranges are the same for all sheets before running this macro. _
    The user wants the *same value/formula* from the first sheet no all other sheets. _
        -- This macro will not increment, sum, or perform any similar modification to the values/formulas as it goes through successive worksheets. _
 _
This macro's usefulness extends to, but is not limited to, cases where the users wishes to edit numerous like-named sheets samely-format, such as inventory tabs, quantity sheets, or financial analysis sheets.


Option Explicit
Sub CRtLS()

'Welcome and guidance message.
MsgBox "This macro will only operate on sheets with like names. You will be given a series of prompts to guide you through the process, with multiple opportunities to cancel if something looks wrong." & vbCr & vbCr & "The following prompt determines if you wish to identify sheets by prefix, suffix, or both.", vbInformation, "Welcome to the CRtLS Macro!"

Dim SearchType As Variant

Do
    SearchType = InputBox("How should we search for sheet names?" & vbCr & vbCr & "1 = Prefix Only" & vbCr & "2 = Suffix Only" & vbCr & "3 = Both Prefix and Suffix", "Prefix, Suffix, or Both?", "1")
Loop Until SearchType > 0 And SearchType <= 3

If MsgBox("1 = Prefix Only" & vbCr & "2 = Suffix Only" & vbCr & "3 = Both Prefix and Suffix" & vbCr & vbCr & "Response interpreted as " & SearchType & ". Is this ok?", vbQuestion + vbOKCancel, "Confirm Input.") = vbCancel Then GoTo Cancelled

'Determined if searching by prefix and/or suffix! =============================

Dim TWB As Workbook 'Shorthand.
Set TWB = ThisWorkbook

Dim Prefix As String, Suffix As String 'To determine which sheets to edit.
Dim SuffixLength As Integer 'Needed when doing worksheet name checks

If SearchType = 1 Or 3 Then 'Values corresponding to Prefix Only and Both will run this.

    'User Input of Prefix
    Prefix = InputBox(Prompt:="Please enter the letter/number/symbol that prefixes sheets you wish to update. Case Sensitive.", _
                      Title:="Input Desired Sheet Prefix")

    If StrPtr(Prefix) = 0 Then GoTo Cancelled 'StrPtr = 0 if user cancels or otherwise escapes out of Inputbox.

    If InStr(TWB.ActiveSheet.Name, Prefix) <> 1 Then
        '!!! Need to create flag or something to modify Cancelled's output
        GoTo Cancelled
    
    End If
    
End If

If SearchType > 1 Then 'Values of 2 or 3 will run this.
    
    'User Input of Suffix
    Suffix = InputBox(Prompt:="Please enter the letter/number/symbol that suffixes sheets you wish to update. Case Sensitive.", _
                      Title:="Input Desired Sheet Suffix")

    If StrPtr(Suffix) = 0 Then GoTo Cancelled 'StrPtr = 0 if user cancels or otherwise escapes out of Inputbox.

    If InStr(TWB.ActiveSheet.Name, Suffix) Then
        '!!! Need to create flag or something to modify Cancelled's output
        GoTo Cancelled
    
    End If
    
End If

Dim FirstPrefixSheetIndex As Integer, FirstPrefixSheetName As String
FirstPrefixSheetIndex = WhereFirstPrefix(TWB, Prefix) 'Determines sheet index from which to copy headers.
FirstPrefixSheetName = TWB.Worksheets(FirstPrefixSheetIndex).Name

Dim StartIndex As Integer

If FirstPrefixSheetIndex <> TWB.ActiveSheet.Index Then 'Checks if first sheet with prefix doesn't equal current sheet's prefix.
'If it does equal, continue as normal.
    StartIndex = ConfirmStart(TWB, Prefix, "") 'Calling this prompts user to confirm they don't want the first Prefix sheet to be where the macro starts.

End If

'Prefix/Suffix and StartIndex successfully determined! ===============================

'In case users have sheets with Prefix at the end of their workbook which are different and/or they don't wish to change. _
One of the limitations of this macro is that it can only be given one start and end point.
Dim EndCheck As String
Dim EndIndex As Integer
Dim EndWS As Worksheet

EndIndex = FindEnd(TWB, Prefix, StartIndex) 'For vbYes, finds last sheet starting from StartIndex.

If Not EndIndex > StartIndex Then 'Prevents errors by not calling further processes if StartIndex is the last sheet. Macro is useless in such a case.
    MsgBox "Current sheet is the end sheet, macro not needed. Cancelling."
    GoTo Cancelled
    
ElseIf MsgBox("Currently, value or formula data from " & FirstPrefixSheetName & " sheet will be copied to all " & Prefix & " sheets. If this is ok, press Yes. If you have a specific sheet after which the macro should stop, press No.", vbYesNo, "Ending Sheet Question") = vbNo Then
    EndIndex = WhereEnd(TWB, Prefix, StartIndex, EndIndex) 'They need it to stop somewhere in particular: WhereEnd finds out where.
    
End If
    
Set EndWS = TWB.Worksheets(EndIndex)
    
'EndIndex successfully determined. ============================================

Dim MatchRge As Range
Set MatchRge = Application.InputBox(Prompt:="Use your mouse to select or use your keyboard to type in a range of cells to copy into other " & Prefix & " sheets.", _
                                    Title:="Input Range of Cells", Type:=8)

'Shorthands
Dim RgeRows As Integer
Dim RgeCols As Integer

RgeRows = MatchRge.Rows.Count
RgeCols = MatchRge.Columns.Count

'Creates an array which has the size of MatchRge
Dim WhatToCopy() As Variant
ReDim WhatToCopy(RgeRows - 1, RgeCols - 1) 'Must subtract because array is indexed from 0 while range is from 1.

'Successfully got MatchRge and sized WhatToCopy. ==============================

Dim i As Integer
Dim j As Integer

'Determine whether to copy values or formulas
Dim ExtractValues As Variant

Dim Count As Integer 'To prevent user going in an infinite loop in the following Do.
Count = 0

Do

If Count > 1 Then
    If MsgBox("Do you wish to cancel the macro?", vbYesNo, "Cancel Macro?") = vbYes Then GoTo Cancelled
End If

If MsgBox("Do you wish to copy values? Press No if you wish to copy its Formulas instead.", vbYesNo, "Confirm Extraction of Values") = vbYes Then
    ExtractValues = True
    Exit Do
    
End If

If MsgBox("Do you wish to copy relative formulas? Press No if you wish to be re-prompted to copy its values.", vbYesNo, "Confirm Extraction of Relative Formulas") = vbYes Then
    ExtractValues = False
    Exit Do
    
Else
    Count = Count + 1

End If

Loop

'Successfully determined if extracting values or formulas. =======================================

'Iterate through MatchRge to plug values/formulas into WhatToCopy array.
For i = 1 To RgeRows
    
    For j = 1 To RgeCols
        
        'Must offset WhatToCopy value by 1 because array is indexed from 0 while range is from 1.
        If ExtractValues = True Then
            WhatToCopy(i - 1, j - 1) = MatchRge(i, j).Address
            'Now make the values of WhatToCopy into formula referencing first sheet cellvalue by adding StartIndex sheet.
            WhatToCopy(i - 1, j - 1) = "='" & ActiveSheet.Name & "'!" & WhatToCopy(i - 1, j - 1)
        
        Else
            WhatToCopy(i - 1, j - 1) = MatchRge(i, j).Formula
            
            If WhatToCopy(i - 1, j - 1) <> "" Then 'Prevents just writing an equals sign into another sheet's cell, just leaves it blank like the parent cell.
                WhatToCopy(i - 1, j - 1) = "=" & WhatToCopy(i - 1, j - 1)
            End If
            
        End If
        
    Next j
    
Next i

'Successfully populated WhatToCopy. ===========================================

Dim k As Integer

'The loop that actually does the editing.
For k = StartIndex + 1 To EndIndex

    If InStr(TWB.Worksheets(k).Name, Prefix) <> 1 Then GoTo NextIteration
    
    For i = 1 To RgeRows
    
        For j = 1 To RgeCols
        
            TWB.Worksheets(k).Range(MatchRge(i, j).Address) = WhatToCopy(i - 1, j - 1)
        
        Next j
        
    Next i
    
NextIteration:
    
Next k

'Successfully input data into indicated sheets. ===============================

Dim SuccessType As String

If ExtractValues = True Then
    SuccessType = "values"

Else
    SuccessType = "relative formulas"

End If

MsgBox Prefix & "-prefixed sheets from " & TWB.Worksheets(StartIndex).Name & " to " & EndWS.Name & " now have matching " & SuccessType & " for range " & MatchRge.Address & ".", , "Success!"
End

'Success! Unless we forcibily cancel or user voluntarily cancels. =============

Cancelled:

MsgBox "Macro Cancelled. No edits made."
End

End Sub
Private Function ConfirmStart(TWB As Workbook, Prefix As String, Flag As String)
    
'Because the sub uses ActiveSheet to set MatchRge, this code attempts to prevents the user _
from changing something from a "wrong" sheet or having to re-run later due to incompleteness.
    
Dim MsgTxt As String
Dim vbType As VbMsgBoxStyle
    
If Flag = "" Then
    MsgTxt = "You are not currently in the first sheet with prefix, " & Prefix & ". If your intention is to run the macro beginning instead at this sheet, hit OK. Otherwise, Cancel and re-call this macro while on the first sheet."
    vbType = vbOKCancel
    
Else
    MsgTxt = "You are not on a sheet with prefix " & Prefix & ". Switch to such a sheet, then re-call this macro."
    vbType = vbOKOnly
    End
        
End If
    
If MsgBox(MsgTxt, vbOKCancel, "Start Sheet Discrepancy") = vbCancel Then
    MsgBox "Macro Cancelled. No edits made." 'Cancels.
    End
        
Else
    ConfirmStart = TWB.ActiveSheet.Index
End If

End Function
Private Function FindEnd(TWB As Workbook, Prefix As String, StartIndex As Integer) As Integer

Dim i As Integer

For i = StartIndex To TWB.Worksheets.Count

    If InStr(TWB.Worksheets(i).Name, Prefix) = 1 Then
        FindEnd = i
    End If
    
Next i

End Function
Private Function WhereEnd(TWB As Workbook, Prefix As String, StartIndex As Integer, EndIndex As Integer) As Integer

Dim i As Long
Dim Confirm As Variant
Dim UserIndication As Variant

Do
    UserIndication = InputBox(Prompt:="Type in exactly the name of the sheet after which you want the macro to stop. Case Sensitive.", _
                              Title:="Give Name of Last Sheet to Edit")

    If StrPtr(UserIndication) = 0 Then GoTo Cancelled 'Cancel if they cancel.
    
    For i = 1 To TWB.Worksheets.Count 'Inner loop begins to find index of sheet.
    
        If TWB.Worksheets(i).Name = UserIndication Then
            
            If Not i > StartIndex Then 'Checks if indicated sheet is not after start sheet, will prompt user again if so.
                MsgBox "Worksheet found, but either equals or comes before " & ThisWorkbook.ActiveSheet.Name & "", vbOKOnly, "Not a Valid Sheet"
                Exit For
            
            Else
                WhereEnd MsgBox("Found sheet " & UserIndication & " with index " & i & " out of " & TWB.Worksheets.Count & " total sheets. Confirm by hitting Yes or hit No to select another sheet name.", vbYesNoCancel, "Confirm Sheet to End")
            
                If WhereEnd = vbYes Or WhereEnd = vbNo Then
                    Exit For
            
                ElseIf WhereEnd = vbCancelled Then GoTo Cancelled 'Cancel if they cancel.'
                       
            End If
            
        ElseIf i = TWB.Worksheets.Count Then 'End of outer If.
            MsgBox "Sheet name not found. This is case sensitive and requires an exact match."
            Exit For
               
        End If
        
    Next i 'Inner loop ends.

Loop Until TWB.Worksheets(i).Name = UserIndication And Confirm = vbYes

WhereEnd = i

Exit Function

Cancelled:
    MsgBox "Macro Cancelled."
    End

End Function
Private Function WhereFirstPrefix(TWB As Workbook, Prefix As String)

Dim i As Long

For i = 1 To TWB.Worksheets.Count
    
    If InStr(TWB.Worksheets(i).Name, Prefix) = 1 Then
        WhereFirstPrefix = i 'Found the first sheet with Prefix prefix.
        Exit Function 'Go back to Sub
    End If

    If i >= TWB.Worksheets.Count And InStr(TWB.Worksheets(i).Name, Prefix) <> 1 Then 'In the case the For Each goes through all sheets but none have the prefix.
        MsgBox "Failed to find any sheet starting with " & Prefix & ", cancelling.", vbCritical
        End
        
    End If
    
Next i

End Function
