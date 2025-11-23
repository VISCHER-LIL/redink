' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict Off

Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Microsoft.Office.Interop.Word
Imports Slib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Public Sub RemoveContentControlsRespectSelection()
        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection

            Dim hasSelection As Boolean = (sel IsNot Nothing AndAlso sel.Range IsNot Nothing AndAlso sel.Range.Start <> sel.Range.End)
            Dim removedCount As Integer = 0

            If hasSelection Then
                removedCount = RemoveContentControlsInRangeKeepContents(sel.Range)
            Else
                Dim Answer As Integer = ShowCustomYesNoBox("No text selection detected. Do you want to remove ALL content controls in the entire document?", "Yes", "No, abort")
                If Answer = 1 Then
                    removedCount = RemoveAllContentControlsKeepContents(app)
                Else
                    ShowCustomMessageBox("Operation aborted.")
                    Exit Sub
                End If
            End If

            ShowCustomMessageBox("Successfully removed " & removedCount.ToString() & " content control(s). Text and formatting were preserved.")
        Catch ex As System.Exception
            ShowCustomMessageBox("Error while removing content controls: " & ex.Message)
        End Try
    End Sub

    Public Function RemoveAllContentControlsKeepContents(app As Microsoft.Office.Interop.Word.Application) As System.Int32
        Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
        If doc Is Nothing Then Return 0

        If doc.ProtectionType <> Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            Throw New System.Exception("Document is protected; cannot remove content controls.")
        End If

        Dim beforeCount As Integer = doc.ContentControls.Count

        ' Snapshot and sort inner-first
        Dim list As New List(Of Microsoft.Office.Interop.Word.ContentControl)(beforeCount)
        For i As Integer = 1 To beforeCount
            list.Add(doc.ContentControls(i))
        Next
        list.Sort(Function(a, b) b.Range.Start.CompareTo(a.Range.Start))

        Dim trackWasOn As Boolean = doc.TrackRevisions
        doc.TrackRevisions = False
        Try
            For Each cc In list
                Try
                    If cc Is Nothing Then Continue For
                    If cc.LockContentControl Then cc.LockContentControl = False
                    If cc.LockContents Then cc.LockContents = False
                    cc.Delete(False) ' keep contents/formatting
                Catch
                    ' continue with other controls
                End Try
            Next
        Finally
            doc.TrackRevisions = trackWasOn
        End Try

        Dim afterCount As Integer = doc.ContentControls.Count
        Return System.Math.Max(0, beforeCount - afterCount)
    End Function



    Public Function RemoveContentControlsInRangeKeepContents(ByVal rng As Microsoft.Office.Interop.Word.Range) As System.Int32
        If rng Is Nothing Then Throw New System.Exception("Selection range is not available.")
        Dim doc As Microsoft.Office.Interop.Word.Document = rng.Document
        If doc Is Nothing Then Return 0

        If doc.ProtectionType <> Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection Then
            Throw New System.Exception("Document is protected; cannot remove content controls.")
        End If

        ' Collect all controls overlapping the selection, same story only
        Dim allCcs As Microsoft.Office.Interop.Word.ContentControls = doc.ContentControls
        Dim list As New List(Of Microsoft.Office.Interop.Word.ContentControl)
        For i As Integer = 1 To allCcs.Count
            Dim cc = allCcs(i)
            If cc.Range Is Nothing Then Continue For
            If cc.Range.StoryType <> rng.StoryType Then Continue For
            If cc.Range.Start < rng.End AndAlso cc.Range.End > rng.Start Then
                list.Add(cc)
            End If
        Next

        If list.Count = 0 Then Return 0

        ' Sort inner-first
        list.Sort(Function(a, b) b.Range.Start.CompareTo(a.Range.Start))

        Dim removed As Integer = 0
        Dim trackWasOn As Boolean = doc.TrackRevisions
        doc.TrackRevisions = False
        Try
            For Each cc In list
                Try
                    If cc Is Nothing Then Continue For
                    If cc.LockContentControl Then cc.LockContentControl = False
                    If cc.LockContents Then cc.LockContents = False
                    cc.Delete(False)
                    removed += 1
                Catch
                    ' ignore and continue
                End Try
            Next
        Finally
            doc.TrackRevisions = trackWasOn
        End Try

        Return removed
    End Function


    Public Async Sub ImportTextFile()
        Dim sel As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        Dim Doc = Await GetFileContent(Nothing, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object))
        sel.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
        sel.Text = Doc
        'sel.End = sel.Start + Doc.Length
        sel.Select()
    End Sub

    Public Sub AcceptFormatting()

        Dim sel As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        Dim formatChangeCount As Integer = 0
        Dim DocRef As String = "in the selected text"

        ' Ensure a selection is made (use content if selection empty)
        If sel Is Nothing OrElse sel.Start = sel.End Then
            sel = Globals.ThisAddIn.Application.ActiveDocument.Content
            DocRef = "in the document"
        End If

        ' Quick exit if no revisions at all
        If sel.Revisions.Count = 0 Then
            ShowCustomMessageBox($"No revisions found {DocRef}. Note: Formatting embedded in insert/delete revisions would also count as those insert/delete changes.")
            Return
        End If

        Dim splash As New Slib.SplashScreen("Accepting formatting-only revisions... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        ' Revision types treated as pure formatting (will be accepted)
        Dim formattingTypes As Word.WdRevisionType() = {
            Word.WdRevisionType.wdRevisionProperty,
            Word.WdRevisionType.wdRevisionParagraphNumber,
            Word.WdRevisionType.wdRevisionParagraphProperty,
            Word.WdRevisionType.wdRevisionSectionProperty,
            Word.WdRevisionType.wdRevisionStyle,
            Word.WdRevisionType.wdRevisionStyleDefinition,
            Word.WdRevisionType.wdRevisionTableProperty
        }

        Dim formattingSet As New HashSet(Of Integer)(formattingTypes.Select(Function(t) CInt(t)))

        ' Structural revision types that may carry embedded formatting the user asked NOT to accept
        Dim structuralTypes As Word.WdRevisionType() = {
            Word.WdRevisionType.wdRevisionInsert,
            Word.WdRevisionType.wdRevisionDelete,
            Word.WdRevisionType.wdRevisionMovedFrom,
            Word.WdRevisionType.wdRevisionMovedTo
        }
        Dim structuralSet As New HashSet(Of Integer)(structuralTypes.Select(Function(t) CInt(t)))

        Dim aborted As Boolean = False

        ' Accept formatting-only revisions
        For Each rev As Word.Revision In sel.Revisions
            System.Windows.Forms.Application.DoEvents()

            If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                aborted = True
                Exit For
            End If
            If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                aborted = True
                Exit For
            End If

            If formattingSet.Contains(CInt(rev.Type)) Then
                Try
                    rev.Accept()
                    formatChangeCount += 1
                Catch
                    ' Ignore failures; continue
                End Try
            End If
        Next

        splash.Close()

        ' Count remaining structural revisions (potentially with embedded formatting)
        Dim embeddedStructuralCount As Integer = sel.Revisions.Cast(Of Word.Revision)().Count(Function(r) structuralSet.Contains(CInt(r.Type)))

        ' Build final message
        Dim msg As New System.Text.StringBuilder
        If aborted Then
            msg.AppendLine("Operation aborted by user (Esc).")
            If formatChangeCount > 0 Then
                msg.AppendLine($"{formatChangeCount} formatting revision(s) were accepted before abort.")
            Else
                msg.AppendLine("No formatting revisions were accepted before abort.")
            End If
        Else
            If formatChangeCount > 0 Then
                msg.AppendLine($"{formatChangeCount} formatting revision(s) {DocRef} (including paragraph numbering) have been accepted.")
            Else
                msg.AppendLine($"No pure formatting revisions were found {DocRef}.")
            End If
        End If

        ' Always inform about possible embedded formatting
        If embeddedStructuralCount > 0 Then
            msg.AppendLine()
            msg.AppendLine($"{embeddedStructuralCount} insertion/deletion/move revision(s) remain. Some formatting applied during those changes cannot be accepted separately.")
        Else
            msg.AppendLine()
            msg.AppendLine("Note: If formatting was applied during text insertions/deletions/moves, it is part of those tracked text changes and cannot be accepted without accepting the text change itself.")
        End If

        ShowCustomMessageBox(msg.ToString())
    End Sub




    Private Shared LastRegexPattern As String = String.Empty
    Private Shared LastRegexOptions As String = String.Empty
    Private Shared LastRegexReplace As String = String.Empty

    Public Sub RegexSearchReplace()
        Dim sel As Word.Range = Globals.ThisAddIn.Application.Selection.Range
        Dim DocRef As String = "in the selected text"

        ' Ensure a selection is made
        If sel Is Nothing OrElse String.IsNullOrWhiteSpace(sel.Text) Then
            Globals.ThisAddIn.Application.ActiveDocument.Content.Select()
            sel = Globals.ThisAddIn.Application.Selection.Range
            DocRef = "in the document"
        End If

        ' Step 1: Get regex patterns
        Dim RegexPattern As String = ShowCustomInputBox("Step 1: Enter your Regex pattern(s), one per line (more info about Regex: vischerlnk.com/regexinfo):", "Regex Search & Replace", False, LastRegexPattern)?.Trim()
        If String.IsNullOrEmpty(RegexPattern) Then Return

        ' Step 2: Get regex options
        Dim optionsInput As String = ShowCustomInputBox("Enter regex option(s) (i for IgnoreCase, m for Multiline, s for Singleline, c for Compiled, r for RightToLeft, e for ExplicitCapture):", "Regex Search & Replace", True, LastRegexOptions)

        Dim regexOptions As RegexOptions = regexOptions.None

        If Not String.IsNullOrEmpty(optionsInput) Then
            ' Add specific options based on user input
            If optionsInput.Contains("i") Then regexOptions = regexOptions Or regexOptions.IgnoreCase
            If optionsInput.Contains("m") Then regexOptions = regexOptions Or regexOptions.Multiline
            If optionsInput.Contains("s") Then regexOptions = regexOptions Or regexOptions.Singleline
            If optionsInput.Contains("c") Then regexOptions = regexOptions Or regexOptions.Compiled
            If optionsInput.Contains("r") Then regexOptions = regexOptions Or regexOptions.RightToLeft
            If optionsInput.Contains("e") Then regexOptions = regexOptions Or regexOptions.ExplicitCapture
        End If

        ' Step 3: Get replacement text
        Dim Replacementtext As String = ShowCustomInputBox("Step 2: Enter your replacement text(s), one on each line, matching to your pattern(s) (leave empty or cancel to only search for the first hit):", "Regex Search & Replace", False, LastRegexReplace)

        ' Update the last-used regex pattern and options
        LastRegexPattern = RegexPattern
        LastRegexOptions = optionsInput
        LastRegexReplace = Replacementtext

        ' Split patterns and replacements into lines
        Dim patterns() As String = RegexPattern.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
        Dim replacements() As String = If(Not String.IsNullOrEmpty(Replacementtext), Replacementtext.Split(New String() {Environment.NewLine}, StringSplitOptions.None), Nothing)

        ' Check if patterns and replacements match
        If replacements IsNot Nothing AndAlso patterns.Length <> replacements.Length Then
            ShowCustomMessageBox("The number of regex patterns does not match the number of replacement lines. Aborting without any replacements done.")
            Return
        End If

        ' Validate all regex patterns first
        For Each pattern As String In patterns
            Try
                Dim regexTest As New Regex(pattern, regexOptions)
            Catch ex As ArgumentException
                ShowCustomMessageBox($"Your regex pattern '{pattern}' is invalid ({ex.Message}). Aborting without any replacements done.")
                Return
            End Try
        Next

        ' Perform replacements after validation
        Dim totalReplacements As Integer = 0

        For i As Integer = 0 To patterns.Length - 1
            Dim pattern As String = patterns(i)
            Dim replacement As String = If(replacements IsNot Nothing, replacements(i), Nothing)

            Dim regex As New Regex(pattern, regexOptions)

            If Not String.IsNullOrEmpty(replacement) Then
                ' Perform replacement
                Dim replacementCount As Integer = 0
                sel.Text = regex.Replace(sel.Text, Function(match)
                                                       replacementCount += 1
                                                       Return replacement
                                                   End Function)
                totalReplacements += replacementCount
            Else
                ' Perform search only
                Dim match As Match = regex.Match(sel.Text)
                If match.Success Then
                    ' Highlight the first match
                    sel.Start = sel.Start + match.Index
                    sel.End = sel.Start + match.Length
                    Globals.ThisAddIn.Application.Selection.Select()
                    Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(sel, True)
                    Return
                Else
                    ShowCustomMessageBox($"No matches found for '{pattern}' {DocRef}.")
                    Return
                End If
            End If
        Next

        If replacements IsNot Nothing Then
            ShowCustomMessageBox($"{totalReplacements} replacement(s) made {DocRef}.")
        Else
            ShowCustomMessageBox("Search complete. No replacements were made.")
        End If
    End Sub

    Public Sub CalculateUserMarkupTimeSpan()

        Try
            Dim userName As String
            Dim docRevisions As Word.Revisions
            Dim rev As Word.Revision
            Dim comment As Word.Comment
            Dim firstTimestamp As Date
            Dim lastTimestamp As Date
            Dim found As Boolean
            Dim userInput As String
            Dim userNames As New Microsoft.VisualBasic.Collection
            Dim selRange As Word.Range
            Dim outputUserNames As String
            Dim DocRef As String = "in the selected text"

            ' Initialize
            found = False
            firstTimestamp = #1/1/1900# ' Default initialization
            lastTimestamp = #1/1/1900# ' Default initialization

            ' Prompt for user input
            userName = Globals.ThisAddIn.Application.UserName

            ' Prompt for user input
            userInput = ShowCustomInputBox("Please enter the name of the user (leave empty for all users):", "Markup Time Span", True, userName)
            userInput = userInput.Trim()


            ' ————————————————————————————————————————————————————————————
            ' Prompt für earliest date
            Dim userDateInput As String
            Dim earliestDate As System.DateTime = System.DateTime.MinValue
            Dim earliestDateFiltered As Boolean = False

            userDateInput = ShowCustomInputBox(
                    "Please enter the earliest date (and time, if you wish) to consider (leave empty for no filter):",
                    "Markup Time Span",
                    True,
                    System.DateTime.Now.AddDays(-2).ToString(System.Globalization.CultureInfo.CurrentCulture)
                )
            userDateInput = userDateInput.Trim()

            Dim parsed As System.DateTime
            If String.IsNullOrEmpty(userDateInput) Then
                earliestDateFiltered = False
            ElseIf System.DateTime.TryParse(
                      userDateInput,
                      System.Globalization.CultureInfo.CurrentCulture,
                      System.Globalization.DateTimeStyles.None,
                      parsed
                  ) Then
                earliestDate = parsed
                earliestDateFiltered = True
            Else
                ShowCustomMessageBox("Improper date/time format - will abort.")
                Exit Sub
            End If

            ' ————————————————————————————————————————————————————————————


            ' Check selection
            If Globals.ThisAddIn.Application.Selection Is Nothing OrElse String.IsNullOrWhiteSpace(Globals.ThisAddIn.Application.Selection.Range.Text) Then
                ' If no selection, select the entire document
                Globals.ThisAddIn.Application.ActiveDocument.Content.Select()
                DocRef = "in the document"
            End If
            selRange = Globals.ThisAddIn.Application.Selection.Range
            docRevisions = selRange.Revisions ' Only consider changes in the selected range

            ' Process revisions
            For Each rev In docRevisions
                If (String.IsNullOrEmpty(userInput) OrElse rev.Author.Equals(userInput, StringComparison.OrdinalIgnoreCase)) _
                       AndAlso (Not earliestDateFiltered OrElse rev.Date >= earliestDate) Then
                    ' Update timestamps
                    If Not found Then
                        firstTimestamp = rev.Date
                        lastTimestamp = rev.Date
                        found = True
                    Else
                        If rev.Date < firstTimestamp Then firstTimestamp = rev.Date
                        If rev.Date > lastTimestamp Then lastTimestamp = rev.Date
                    End If
                    ' Collect user names if processing all
                    Try
                        userNames.Add(rev.Author, rev.Author.ToLower())
                    Catch ex As Exception
                        ' Ignore duplicates
                    End Try
                End If
            Next

            ' Process comments
            For Each comment In selRange.Comments
                If (String.IsNullOrEmpty(userInput) OrElse comment.Author.Equals(userInput, StringComparison.OrdinalIgnoreCase)) _
                       AndAlso (Not earliestDateFiltered OrElse comment.Date >= earliestDate) Then

                    ' Update timestamps
                    If Not found Then
                        firstTimestamp = comment.Date
                        lastTimestamp = comment.Date
                        found = True
                    Else
                        If comment.Date < firstTimestamp Then firstTimestamp = comment.Date
                        If comment.Date > lastTimestamp Then lastTimestamp = comment.Date
                    End If
                    ' Collect user names if processing all
                    Try
                        userNames.Add(comment.Author, comment.Author.ToLower())
                    Catch ex As Exception
                        ' Ignore duplicates
                    End Try
                End If
            Next

            ' Display results
            If found Then
                Dim timeSpan As String
                Dim timeDiff As Double
                timeDiff = DateDiff(DateInterval.Minute, firstTimestamp, lastTimestamp) ' Time difference in minutes
                timeSpan = System.Math.Floor(timeDiff / 1440).ToString() & " days, " &
                       ((timeDiff Mod 1440) \ 60).ToString("00") & " hours, " &
                       (timeDiff Mod 60).ToString("00") & " minutes"

                ' Format timestamps without seconds
                Dim formattedFirstTimestamp As String
                Dim formattedLastTimestamp As String
                formattedFirstTimestamp = firstTimestamp.ToString("dd/MM/yyyy HH:mm")
                formattedLastTimestamp = lastTimestamp.ToString("dd/MM/yyyy HH:mm")
                If String.IsNullOrEmpty(userInput) Then
                    ' Display all users
                    Dim user As Object
                    outputUserNames = "Users involved:" & vbCrLf
                    For Each user In userNames
                        outputUserNames &= "- " & user.ToString() & vbCrLf
                    Next
                Else
                    outputUserNames = "User: " & userInput
                End If
                ShowCustomMessageBox(outputUserNames & vbCrLf & If(earliestDateFiltered, "Earliest considered: " & earliestDate.ToString("dd/MM/yyyy HH:mm") & vbCrLf, "") & "First markup/comment: " & formattedFirstTimestamp & vbCrLf &
    "Last markup/comment: " & formattedLastTimestamp & vbCrLf &
    "Time span: " & timeSpan)
            Else
                If String.IsNullOrEmpty(userInput) Then
                    ShowCustomMessageBox($"No markups or comments found {DocRef}.")
                Else
                    ShowCustomMessageBox("No markups or comments found for user '" & userInput & $"' {DocRef}.")
                End If
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in CalculateUserMarkupTimeSpan: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub CompareSelectionHalves()

        Dim sel As Word.Range
        Dim nonEmptyParaCount As Long
        Dim halfParaCount As Long
        Dim firstRange As Word.Range
        Dim secondRange As Word.Range
        Dim paraIndices() As Long
        Dim i As Long, index As Long

        ' Get the selected text
        sel = Globals.ThisAddIn.Application.Selection.Range

        ' Count non-empty paragraphs and store their indices
        ReDim paraIndices(0 To sel.Paragraphs.Count - 1)
        index = 0
        For i = 1 To sel.Paragraphs.Count
            If Len(sel.Paragraphs(i).Range.Text.Trim()) > 1 Then ' Greater than 1 to account for paragraph mark
                index += 1
                paraIndices(index - 1) = i
            End If
        Next

        ' Update nonEmptyParaCount
        nonEmptyParaCount = index

        ' If number of non-empty paragraphs is uneven or zero, abort
        If nonEmptyParaCount Mod 2 <> 0 Or nonEmptyParaCount = 0 Then
            ShowCustomMessageBox("The number of non-empty paragraphs in the selection is uneven or zero. Please select an even number of non-empty paragraphs.")
            Return
        End If

        ' Determine the halfway point
        halfParaCount = nonEmptyParaCount \ 2

        ' Get the first half and second half ranges
        firstRange = sel.Paragraphs(paraIndices(0)).Range
        firstRange.End = sel.Paragraphs(paraIndices(halfParaCount - 1)).Range.End

        secondRange = sel.Paragraphs(paraIndices(halfParaCount)).Range
        secondRange.End = sel.Paragraphs(paraIndices(nonEmptyParaCount - 1)).Range.End


        ' Get text from the first and second range without the final paragraph marks
        Dim text1 As String = Left(firstRange.Text, Len(firstRange.Text) - 1)
        Dim text2 As String = Left(secondRange.Text, Len(secondRange.Text) - 1)

        If INI_MarkupMethodHelper <> 1 Then
            CompareAndInsert(text1, text2, secondRange, INI_MarkupMethodHelper = 3, "These are the differences of the second (set of) paragraph(s) of the text selected:")
        Else
            CompareAndInsertComparedoc(text1, text2, secondRange)
        End If
    End Sub

End Class
