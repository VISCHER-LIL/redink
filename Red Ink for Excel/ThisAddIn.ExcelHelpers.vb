' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.ExcelHelpers.vb
' Purpose: Excel helper routines for row height adjustment (including merged cells),
'          legacy comment (note) shape sizing, and multi-pattern regex search/replace
'          across a selected range or entire worksheet.
'
' Architecture:
' - Operates on the active worksheet (Globals.ThisAddIn.Application.ActiveSheet).
' - Each routine acquires the current selection; if empty, user may opt to use UsedRange.
' - SplashScreen shown; ESC key aborts loops (checked via GetAsyncKeyState).
' - AdjustHeight: AutoFits rows, measures required heights (merged cells handled by temporary unmerge and width aggregation), tracks original/max heights, applies final capped height (<= 409).
' - AdjustLegacyNotes: Resizes legacy Comment shapes; constrains width (70–250) and computes height from text length and font size approximation.
' - RegexSearchReplace: Collects multi-line regex patterns and optional replacements, validates all patterns, applies ordered replacements to string cells, counts modifications.
' - Error handling: Each method catches System.Exception and reports via MessageBox.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Hosts helper routines used by the Excel add-in for row sizing, legacy note sizing, and regex-based edits.
''' </summary>
Partial Public Class ThisAddIn

    ''' <summary>
    ''' Adjusts row heights for the selected cell range (or entire UsedRange if approved), handling merged cells,
    ''' preserving original heights when larger, and capping at 409 points. ESC aborts processing.
    ''' </summary>
    ''' <param name="Silent">True to suppress prompt when selection is empty; False to ask user.</param>
    ''' <remarks>
    ''' Uses AutoFit, forces WrapText for height calculation, temporarily unmerges horizontally merged cells to aggregate column widths,
    ''' restores widths, then re-merges. Tracks original and maximum computed heights per row.
    ''' </remarks>
    Public Sub AdjustHeight(Optional Silent As Boolean = False)

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active Excel worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the current selection
            Dim selectedRange As Excel.Range = CType(Globals.ThisAddIn.Application.Selection, Excel.Range)
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' Check if the selection is empty or null
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As Integer = 0
                If Not Silent Then
                    result = ShowCustomYesNoBox("No cells are selected. Would you like to perform the operation on the entire worksheet?", "Yes", "No", "Adjust Height")
                End If
                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                Else
                    If Not Silent Then ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Perform AutoFit on the rows of the selected range to ensure initial proper height
            selectedRange.Rows.AutoFit()

            ' Prepare dictionaries for tracking row heights
            Dim rowOriginalHeights As New Dictionary(Of Integer, Double)()
            Dim rowMaxHeights As New Dictionary(Of Integer, Double)()

            ' Initialize these dictionaries for each row in the selection
            For Each oneRow As Excel.Range In selectedRange.Rows
                Dim rowIndex As Integer = oneRow.Row
                Dim currentHeight As Double = CDbl(CType(activeSheet.Rows(rowIndex), Excel.Range).RowHeight)
                rowOriginalHeights(rowIndex) = currentHeight
                ' Start the max at whatever the row is currently
                rowMaxHeights(rowIndex) = currentHeight
            Next

            splash.Show()
            splash.Refresh()

            ' Iterate through each cell in the selection
            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                If cell Is Nothing Then Continue For

                ' We'll always enable wrapping so AutoFit will compute multi-line height
                cell.WrapText = True

                Dim wasMerged As Boolean = CBool(cell.MergeCells)
                Dim mergeArea As Excel.Range = If(wasMerged, cell.MergeArea, cell)

                ' Temporarily store the row index for dictionary look-up
                Dim rowIndex As Integer = mergeArea.Row

                ' We'll measure how tall Excel wants to make this cell
                Dim newHeight As Double = 0

                If wasMerged Then
                    ' Store the original column widths for each column
                    Dim firstColIndex As Integer = mergeArea.Column
                    Dim totalCols As Integer = mergeArea.Columns.Count
                    Dim originalWidths As New List(Of Double)

                    For iCol As Integer = 0 To totalCols - 1
                        Dim colWidth As Double = CDbl(CType(activeSheet.Columns(firstColIndex + iCol), Excel.Range).ColumnWidth)
                        originalWidths.Add(colWidth)
                    Next

                    ' Sum the widths so we can set it on the first column after unmerging
                    Dim combinedWidth As Double = originalWidths.Sum()

                    ' Unmerge
                    mergeArea.UnMerge()

                    ' Set only the first column to the combined width so AutoFit sees the "true" width
                    CType(activeSheet.Columns(firstColIndex), Excel.Range).ColumnWidth = combinedWidth

                    ' Autofit (note: must do autofit on entire row(s) that the cell spans)
                    mergeArea.Rows.AutoFit()

                    ' Capture the new row height - handle DBNull for vertically merged cells
                    Dim rowHeightValue As Object = mergeArea.RowHeight
                    If rowHeightValue IsNot Nothing AndAlso Not IsDBNull(rowHeightValue) Then
                        newHeight = CDbl(rowHeightValue)
                    Else
                        ' For vertically merged cells, get height from first row
                        Dim firstRow As Excel.Range = CType(mergeArea.Rows(1), Excel.Range)
                        newHeight = CDbl(firstRow.RowHeight)
                    End If

                    ' Restore the original column widths
                    For iCol As Integer = 0 To totalCols - 1
                        CType(activeSheet.Columns(firstColIndex + iCol), Excel.Range).ColumnWidth = originalWidths(iCol)
                    Next

                    ' Re-merge
                    Dim remergeRange As Excel.Range = CType(activeSheet.Range(
                        CType(activeSheet.Cells(mergeArea.Row, firstColIndex), Excel.Range),
                        CType(activeSheet.Cells(mergeArea.Row, firstColIndex + totalCols - 1), Excel.Range)
                    ), Excel.Range)
                    remergeRange.Merge()

                Else
                    ' If not merged, simply use AutoFit
                    mergeArea.Rows.AutoFit()
                    Dim rowHeightValue As Object = mergeArea.RowHeight
                    If rowHeightValue IsNot Nothing AndAlso Not IsDBNull(rowHeightValue) Then
                        newHeight = CDbl(rowHeightValue)
                    End If
                End If

                ' Store the maximum needed height for this row so far
                If rowMaxHeights.ContainsKey(rowIndex) Then
                    ' Compare existing max with newly measured height
                    If newHeight > rowMaxHeights(rowIndex) Then
                        rowMaxHeights(rowIndex) = newHeight
                    End If
                End If

            Next

            ' Now set each row’s height to the maximum of:
            For Each rowIndex As Integer In rowMaxHeights.Keys.ToList()

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                Dim finalHeight As Double = Math.Max(rowMaxHeights(rowIndex), rowOriginalHeights(rowIndex))
                If finalHeight > 409 Then finalHeight = 409

                CType(activeSheet.Rows(rowIndex), Excel.Range).RowHeight = finalHeight
            Next

        Catch ex As System.Exception
            MessageBox.Show($"Error in AdjustHeight: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try

    End Sub

    ''' <summary>
    ''' Resizes legacy comment (note) shapes in the selected range (or UsedRange if chosen) by constraining width
    ''' and computing required height from text length and font size. ESC aborts processing.
    ''' </summary>
    ''' <remarks>
    ''' Width constrained to 70–250 points; height based on AutoSize minimum and an estimated line height.
    ''' </remarks>
    Public Sub AdjustLegacyNotes()

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active Excel worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the current selection
            Dim selectedRange As Excel.Range = CType(Globals.ThisAddIn.Application.Selection, Excel.Range)
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' Check if the selection is empty or null
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As Integer = ShowCustomYesNoBox(
                    "No cells are selected. Would you like to perform the operation on the entire worksheet?",
                    "Yes",
                    "No",
                    "Adjust Legacy Notes"
                )

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                Else
                    ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Perform AutoFit on the rows of the selected range to ensure initial proper height
            selectedRange.Rows.AutoFit()

            splash.Show()
            splash.Refresh()

            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                If cell Is Nothing Then Continue For

                If cell.Comment IsNot Nothing Then

                    ' Ensure the note box dimensions are at least 70 wide and 20 high, and no more than 250 wide
                    Dim comment As Excel.Comment = cell.Comment
                    With comment.Shape

                        .TextFrame.AutoSize = True
                        Dim MinimumHeight As Double = .Height

                        .TextFrame.AutoSize = False

                        ' Enforce width constraints
                        If .Width < 70 Then .Width = 70
                        If .Width > 250 Then .Width = 250

                        ' Dynamically calculate and set height
                        Dim textLength As Integer = Len(comment.Text)
                        Dim fontSize As Double = CDbl(.TextFrame.Characters.Font.Size)
                        Dim lines As Integer = CInt(Math.Ceiling(textLength / (250 / (fontSize - 2)))) ' Approximation based on average char width
                        Dim lineHeight As Double = fontSize + 2 ' Approximate height per line in points
                        Dim requiredHeight As Double = Math.Max(MinimumHeight, (lines * lineHeight)) + 10

                        If lines > 1 Then .Width = 250

                        .Height = CSng(requiredHeight)

                    End With
                End If

            Next

        Catch ex As System.Exception
            MessageBox.Show($"Error in AdjustLegacyNotes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try

    End Sub

    ''' <summary>
    ''' Stores the last entered regex pattern list (multi-line, one pattern per line).
    ''' </summary>
    Private Shared LastRegexPattern As String = String.Empty  ' Last entered pattern(s) (multi-line).

    ''' <summary>
    ''' Stores the last entered regex option flags for reuse.
    ''' </summary>
    Private Shared LastRegexOptions As String = String.Empty  ' Last entered option flags.

    ''' <summary>
    ''' Stores the last entered replacement text lines aligned with patterns.
    ''' </summary>
    Private Shared LastRegexReplace As String = String.Empty  ' Last entered replacement line(s).

    ''' <summary>
    ''' Applies one or more regular expression search/replace operations across the selected range (or UsedRange if approved).
    ''' Prompts for patterns, options, and replacements; validates all patterns before applying. ESC aborts processing.
    ''' </summary>
    ''' <remarks>
    ''' Patterns entered line-by-line; replacements must match count (if provided). Counts each cell text modified at least once.
    ''' Supports options: i (IgnoreCase), m (Multiline), s (Singleline), c (Compiled), r (RightToLeft), e (ExplicitCapture).
    ''' </remarks>
    Public Sub RegexSearchReplace()

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the selected range
            Dim selectedRange As Excel.Range = CType(Globals.ThisAddIn.Application.Selection, Excel.Range)
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            Dim processEntireSheet As Boolean = False

            ' If no range is selected, ask to process the entire worksheet
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then

                Dim result As Integer = ShowCustomYesNoBox("No cells are selected. Would you like to perform the operation on the entire worksheet?", "Yes", "No", "Regex Search & Replace")

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                    processEntireSheet = True
                Else
                    ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Step 1: Get regex patterns
            Dim regexPattern As String = ShowCustomInputBox("Step 1: Enter your Regex pattern(s), one per line (more info about Regex: vischerlnk.com/regexinfo):", "Regex Search & Replace", False, LastRegexPattern)?.Trim()
            If String.IsNullOrEmpty(regexPattern) Then Exit Sub

            ' Step 2: Get regex options
            Dim optionsInput As String = ShowCustomInputBox("Enter regex option(s) (i for IgnoreCase, m for Multiline, s for Singleline, c for Compiled, r for RightToLeft, e for ExplicitCapture):", "Regex Search & Replace", True, LastRegexOptions)

            Dim regexOptions As RegexOptions = RegexOptions.None

            If Not String.IsNullOrEmpty(optionsInput) Then
                If optionsInput.Contains("i") Then regexOptions = regexOptions Or RegexOptions.IgnoreCase
                If optionsInput.Contains("m") Then regexOptions = regexOptions Or RegexOptions.Multiline
                If optionsInput.Contains("s") Then regexOptions = regexOptions Or RegexOptions.Singleline
                If optionsInput.Contains("c") Then regexOptions = regexOptions Or RegexOptions.Compiled
                If optionsInput.Contains("r") Then regexOptions = regexOptions Or RegexOptions.RightToLeft
                If optionsInput.Contains("e") Then regexOptions = regexOptions Or RegexOptions.ExplicitCapture
            End If

            ' Step 3: Get replacement text
            Dim replacementText As String = ShowCustomInputBox("Step 2: Enter your replacement text(s), one on each line, matching to your pattern(s):", "Regex Search & Replace", False, LastRegexReplace)

            ' Update the last-used regex pattern and options
            LastRegexPattern = regexPattern
            LastRegexOptions = optionsInput
            LastRegexReplace = replacementText

            ' Split patterns and replacements into lines
            Dim patterns() As String = regexPattern.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            Dim replacements() As String = If(Not String.IsNullOrEmpty(replacementText), replacementText.Split(New String() {Environment.NewLine}, StringSplitOptions.None), Nothing)

            ' Check if patterns and replacements match
            If replacements IsNot Nothing AndAlso patterns.Length <> replacements.Length Then
                ShowCustomMessageBox("The number of regex patterns does not match the number of replacement lines. Aborting without any replacements done.")
                Exit Sub
            End If

            ' Validate all regex patterns first
            For Each pattern As String In patterns
                Try
                    Dim regexTest As New Regex(pattern, regexOptions)
                Catch ex As ArgumentException
                    ShowCustomMessageBox($"Your regex pattern '{pattern}' is invalid ({ex.Message}). Aborting without any replacements done.")
                    Exit Sub
                End Try
            Next

            splash.Show()
            splash.Refresh()

            ' Perform replacements
            Dim totalReplacements As Integer = 0

            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                If cell.Value2 IsNot Nothing AndAlso TypeOf cell.Value2 Is String Then
                    Dim cellText As String = cell.Value2.ToString()

                    For i As Integer = 0 To patterns.Length - 1
                        Dim regex As New Regex(patterns(i), regexOptions)
                        Dim replacement As String = If(replacements IsNot Nothing, replacements(i), Nothing)

                        ' Perform replacement
                        Dim newText As String = regex.Replace(cellText, replacement)
                        If newText <> cellText Then
                            totalReplacements += 1
                            cell.Value2 = newText
                        End If
                    Next
                End If
            Next

            ShowCustomMessageBox($"{totalReplacements} replacement(s) made in the selected cells.")

        Catch ex As System.Exception
            MessageBox.Show($"Error in RegexSearchReplace: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try
    End Sub
End Class