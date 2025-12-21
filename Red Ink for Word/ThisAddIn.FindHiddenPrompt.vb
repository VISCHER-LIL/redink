' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.FindHiddenPrompt.vb
' Purpose: Detects and reveals hidden, obfuscated, or camouflaged text in Word documents
'          that could be used for prompt injection or other malicious purposes.
'
' Architecture:
'  - Multi-Story Support: Scans all document stories (main text, headers, footers, footnotes,
'    endnotes, text frames) to ensure comprehensive coverage.
'  - Two-Pass Detection:
'    1. Deterministic formatting heuristics (hidden text, tiny fonts, white-on-white, etc.)
'    2. LLM-based semantic analysis to detect suspicious content patterns
'  - Progress Tracking: Uses ProgressBarModule for user feedback during long operations;
'    supports cancellation at any point.
'  - Formatting Heuristics (10 checks):
'    * Word Hidden text (Font.Hidden) - revealed in red
'    * Very small font size (< 3pt)
'    * Font color matching background shading
'    * White-on-white text (explicit and Auto-resolved)
'    * Font color matching highlight color
'    * Extreme font scaling (<= 10%)
'    * Negative character spacing (< -2pt)
'    * Font color matching table cell shading
'    * White-on-white in table cells
'    * Zero-width and Bidi control characters
'    * Field codes with formatting switches (MERGEFORMAT, CHARFORMAT)
'  - Story-Aware Processing: Differentiates between commentable stories (main text, text frames)
'    where bubble comments can be added directly, and non-commentable stories (footnotes, endnotes,
'    headers, footers) where findings are aggregated into summary notices.
'  - File Import: Optionally analyzes external files (.txt, .rtf, .doc, .docx, .pdf, .pptx)
'    with OCR support for scanned PDFs.
'  - Result Presentation: Findings displayed as Word comments (bubbles) with prefix "-FHP";
'    footnote/endnote findings shown in summary comment at document end.
'  - External Dependencies: SharedLibrary.SharedMethods for LLM calls, UI dialogs, file I/O,
'    and PDF/PowerPoint processing; ProgressBarModule for cancellable progress tracking.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Main entry point for finding hidden prompts/malicious text in Word documents.
    ''' Supports three modes: (1) selected text only, (2) entire document including all stories,
    ''' (3) imported external file. Performs both deterministic formatting checks and LLM-based
    ''' semantic analysis. Results are presented as Word comments with prefix "-FHP".
    ''' </summary>
    ''' <returns>Task representing the asynchronous operation.</returns>
    ''' <remarks>
    ''' User is prompted to select scope when no text is selected. "Check all text" iterates
    ''' all document stories (main, headers, footers, footnotes, endnotes, text frames).
    ''' "Check file" imports content from external file and analyzes in new document.
    ''' Progress is tracked via ProgressBarModule; operation can be cancelled at any time.
    ''' Hidden text spans are revealed (unhidden and colored red) during detection.
    ''' </remarks>
    Public Async Function FindHiddenPrompts() As System.Threading.Tasks.Task

        If INILoadFail() Then Return

        Dim Prefix As String = "-FHP"

        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection

            Dim JumpRoundA As Boolean = False
            Dim CheckAll As Boolean = False

            Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.")
                Return
            End If

            ' Prompt user for scope if no text selected
            If sel.Type = WdSelectionType.wdSelectionIP Then
                Dim answer As Integer = ShowCustomYesNoBox("You have not selected any text. Check the all text (including foot- and endnotes) or instead check a file?", "Check all text", "Check file")
                If answer = 0 Then Return
                If answer = 1 Then
                    ' Mode: Check all text in active document
                    app.Selection.WholeStory()
                    sel = app.Selection
                    CheckAll = True
                Else
                    ' Mode: Import and check external file
                    DragDropFormLabel = "Document files (.txt, .docx, .pdf) or Powerpoint (.pptx)."
                    DragDropFormFilter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.pptx|" &
                                     "Text Files (*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                                     "Rich Text Files (*.rtf)|*.rtf|" &
                                     "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                                     "PDF Files (*.pdf)|*.pdf|" &
                                     "Powerpoint Files (*.pptx)|*.pptx"

                    Dim FilePath As String = GetFileName()
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    If String.IsNullOrWhiteSpace(FilePath) Then
                        ShowCustomMessageBox("No file has been selected - will abort.")
                        Return
                    End If

                    Dim ext As String = IO.Path.GetExtension(FilePath).ToLowerInvariant()

                    ' Load file content based on extension
                    Dim FromFile As String = ""
                    Select Case ext
                        Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm"
                            FromFile = ReadTextFile(FilePath, True)
                        Case ".rtf"
                            FromFile = ReadRtfAsText(FilePath, True)
                        Case ".doc", ".docx"
                            FromFile = ReadWordDocument(FilePath, True)
                        Case ".pdf"
                            Dim OCRAnswer As Integer = ShowCustomYesNoBox("Do you want to enable OCR for scanned PDFs? This may take longer and not find invisible text (you will be asked to confirm).", "No, proceed without", "Yes, do OCR if needed")
                            If OCRAnswer = 0 Then
                                ShowCustomMessageBox("Aborted by you.")
                                Return
                            End If
                            FromFile = Await ReadPdfAsText(FilePath, True, OCRAnswer = 2, True, _context)
                        Case ".pptx"
                            FromFile = GetPresentationJson(FilePath)
                        Case Else
                            FromFile = "Error: File type not supported."
                    End Select
                    If FromFile.StartsWith("Error:") Then
                        ShowCustomMessageBox(FromFile)
                        Return
                    End If
                    If String.IsNullOrWhiteSpace(FromFile) Then
                        ShowCustomMessageBox("The file you provided did not contain any text - will abort.")
                        Return
                    End If

                    ' Create new document with imported content
                    Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                    newDoc.Activate()

                    Dim rng As Word.Range = newDoc.Content
                    rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Dim safeText As String = FromFile.Replace(ChrW(0), String.Empty)
                    rng.InsertAfter(safeText)

                    newDoc.Content.Select()

                    sel = app.Selection
                    JumpRoundA = True
                End If
            End If

            ' ===== Check all text mode: iterate all document stories =====
            If CheckAll AndAlso Not JumpRoundA Then
                Dim originalStart As Integer = sel.Start
                Dim originalEnd As Integer = sel.End

                ' Build flat list of all story ranges (including linked header/footer sections)
                Dim storyList As New List(Of (rng As Word.Range, st As Word.WdStoryType))()
                For Each firstStory As Word.Range In doc.StoryRanges
                    Dim cur As Word.Range = firstStory
                    Do While cur IsNot Nothing
                        If cur.StoryLength > 0 Then
                            storyList.Add((cur.Duplicate, cur.StoryType))
                        End If
                        cur = cur.NextStoryRange
                    Loop
                Next

                ' Aggregate findings for non-commentable stories (footnotes/endnotes)
                Dim footnoteLines As New List(Of String)()
                Dim endnoteLines As New List(Of String)()

                ' Pass 1: Deterministic formatting checks across ALL stories
                For Each s In storyList
                    Dim tr As Word.Range = s.rng
                    Dim st As Word.WdStoryType = s.st
                    If tr Is Nothing OrElse String.IsNullOrWhiteSpace(tr.Text) Then Continue For

                    Dim roundA As String = BuildSuspicionBubbleString(tr)
                    Debug.WriteLine($"FindHiddenPrompts[RoundA][{st}]: {roundA}")

                    If Not String.IsNullOrEmpty(roundA) Then
                        If IsFootnoteOrEndnote(st) Then
                            ' Aggregate for summary notice (cannot add comments directly)
                            If st = Word.WdStoryType.wdFootnotesStory Then
                                footnoteLines.AddRange(ConvertSetBubblesToLines(roundA))
                            ElseIf st = Word.WdStoryType.wdEndnotesStory Then
                                endnoteLines.AddRange(ConvertSetBubblesToLines(roundA))
                            End If
                        ElseIf IsCommentableStory(st) Then
                            ' Add bubble comments directly to commentable stories
                            Try
                                tr.Select()
                                SetBubbles(roundA, app.Selection, True, Prefix)
                            Catch
                                ' Silently skip if comments fail for this story
                            End Try
                        End If
                    End If
                Next

                ' Pass 2: LLM-based semantic checks across ALL stories
                ShowCustomMessageBox("Now having your LLM check the document (including foot- and endnotes) for potentially malicious content...", autoCloseSeconds:=5, Defaulttext:="")
                Dim systemPrompt As String = InterpolateAtRuntime(SP_FindPrompts & " " & SP_Add_Bubbles)

                For Each s In storyList
                    Dim tr As Word.Range = s.rng
                    Dim st As Word.WdStoryType = s.st
                    If tr Is Nothing OrElse String.IsNullOrWhiteSpace(tr.Text) Then Continue For

                    Dim userPrompt As String = "<TEXTTOPROCESS>" & tr.Text & "</TEXTTOPROCESS>"
                    Dim llmResult As String = Await LLM(systemPrompt, userPrompt, "", "", 0, False)
                    llmResult = If(llmResult, "").Trim()

                    If Not String.IsNullOrEmpty(llmResult) Then
                        If IsFootnoteOrEndnote(st) Then
                            ' Aggregate for summary notice
                            If st = Word.WdStoryType.wdFootnotesStory Then
                                footnoteLines.AddRange(ConvertSetBubblesToLines(llmResult))
                            ElseIf st = Word.WdStoryType.wdEndnotesStory Then
                                endnoteLines.AddRange(ConvertSetBubblesToLines(llmResult))
                            End If
                        ElseIf IsCommentableStory(st) Then
                            ' Add bubble comments directly
                            Try
                                tr.Select()
                                SetBubbles(llmResult, app.Selection, True, Prefix)
                            Catch
                                ' Silently skip if comments fail for this story
                            End Try
                        End If
                    End If
                Next

                ' Add summary comments for footnote/endnote findings at end of main story
                Dim mainEnd As Integer = doc.StoryRanges(Word.WdStoryType.wdMainTextStory).End

                If footnoteLines.Count > 0 Then
                    Dim notice As String = BuildNoticeText("Footnote", footnoteLines)
                    Debug.WriteLine("Footnote Notice: " & notice)
                    LegacyAddNoticeBubbleAt(doc, mainEnd, notice, Prefix)
                End If

                If endnoteLines.Count > 0 Then
                    Dim notice As String = BuildNoticeText("Endnote", endnoteLines)
                    Debug.WriteLine("Endnote Notice: " & notice)
                    LegacyAddNoticeBubbleAt(doc, mainEnd, notice, Prefix)
                End If

                ' Restore original selection
                Try
                    app.Selection.SetRange(originalStart, originalEnd)
                Catch
                End Try

                ShowCustomMessageBox("Analysis completed. See bubble comments and the final notice for footnote/endnote results.")
                Return
            End If

            ' ===== Selection-based or imported file mode =====

            Dim selStart As System.Int32 = sel.Start
            Dim selEnd As System.Int32 = sel.End
            Dim sameSel As Microsoft.Office.Interop.Word.Selection = Nothing

            If Not JumpRoundA Then
                ' Pass 1: Deterministic formatting checks on selection
                Dim roundA As System.String = BuildSuspicionBubbleString(sel.Range)
                Debug.WriteLine("FindHiddenPrompts: Formatting-based findings: " & roundA)

                If Not System.String.IsNullOrEmpty(roundA) Then
                    SetBubbles(roundA, sel, True, Prefix)
                End If

                ' Restore original selection for LLM check
                app.Selection.SetRange(selStart, selEnd)
                sameSel = app.Selection
                If sameSel Is Nothing OrElse sameSel.Range Is Nothing OrElse sameSel.Range.Text Is Nothing _
               OrElse sameSel.Range.Text.Length = 0 Then
                    Return
                End If

                ShowCustomMessageBox("Now having your LLM check the selected text for potentially malicious content...", autoCloseSeconds:=5, Defaulttext:="")
            Else
                sameSel = sel
            End If

            ' Pass 2: LLM-based semantic check
            Dim systemPrompt2 As String = InterpolateAtRuntime(SP_FindPrompts & " " & SP_Add_Bubbles)
            Dim userPrompt2 As String = "<TEXTTOPROCESS>" & sameSel.Range.Text & "</TEXTTOPROCESS>"

            Dim llmResult2 As String = Await LLM(systemPrompt2, userPrompt2, "", "", 0, False)
            llmResult2 = If(llmResult2, "").Trim()
            If llmResult2 Is Nothing OrElse llmResult2.Length = 0 Then
                If JumpRoundA Then
                    ShowCustomMessageBox("No potentially malicious text found.")
                    Return
                End If
            End If

            SetBubbles(llmResult2, sameSel, True, Prefix)
            ShowCustomMessageBox("Analysis completed. See bubble comments for the results.")

        Catch ex As System.Exception
            ShowCustomMessageBox("An error occurred in FindHiddenPrompts: " & ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Determines whether the specified story type is a footnote or endnote story.
    ''' </summary>
    ''' <param name="st">Story type to check.</param>
    ''' <returns>True if story is footnotes or endnotes; otherwise False.</returns>
    Private Function IsFootnoteOrEndnote(st As Word.WdStoryType) As Boolean
        Return st = Word.WdStoryType.wdFootnotesStory OrElse st = Word.WdStoryType.wdEndnotesStory
    End Function

    ''' <summary>
    ''' Determines whether Word comments can be added to the specified story type.
    ''' </summary>
    ''' <param name="st">Story type to check.</param>
    ''' <returns>True if story supports comments (main text, text frames); otherwise False.</returns>
    ''' <remarks>
    ''' Word does not allow comments in headers, footers, footnotes, or endnotes.
    ''' Findings in non-commentable stories must be aggregated into summary notices.
    ''' </remarks>
    Private Function IsCommentableStory(st As Word.WdStoryType) As Boolean
        Return st = Word.WdStoryType.wdMainTextStory OrElse st = Word.WdStoryType.wdTextFrameStory
    End Function

    ''' <summary>
    ''' Converts SetBubbles-format string (text@@comment§§§...) into list of human-readable lines.
    ''' </summary>
    ''' <param name="raw">SetBubbles format string with §§§ record separator and @@ field separator.</param>
    ''' <returns>List of formatted lines: "text": comment</returns>
    ''' <remarks>
    ''' Used to convert findings from non-commentable stories into summary notice format.
    ''' Empty or whitespace-only records are skipped.
    ''' </remarks>
    Private Function ConvertSetBubblesToLines(raw As String) As List(Of String)
        Dim lines As New List(Of String)()
        If String.IsNullOrWhiteSpace(raw) Then Return lines

        Dim recs = raw.Split(New String() {"§§§"}, StringSplitOptions.RemoveEmptyEntries)
        For Each rec In recs
            Dim parts = rec.Split(New String() {"@@"}, 2, StringSplitOptions.None)
            Dim txt As String = If(parts.Length > 0, parts(0), String.Empty)
            Dim cmt As String = If(parts.Length > 1, parts(1), String.Empty)
            If Not String.IsNullOrWhiteSpace(txt) OrElse Not String.IsNullOrWhiteSpace(cmt) Then
                lines.Add($"""{txt}"": {cmt}")
            End If
        Next
        Return lines
    End Function

    ''' <summary>
    ''' Builds formatted notice text for footnote or endnote findings.
    ''' </summary>
    ''' <param name="scope">Story scope name ("Footnote" or "Endnote").</param>
    ''' <param name="lines">List of finding lines to include in notice.</param>
    ''' <returns>Multi-line notice string with heading and findings.</returns>
    ''' <remarks>
    ''' Format: "Suspicious [scope] text:" followed by one finding per line.
    ''' Used to create summary comment at document end for non-commentable stories.
    ''' </remarks>
    Private Function BuildNoticeText(scope As String, lines As List(Of String)) As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($"Suspicious {scope} text:")
        For Each l In lines
            sb.AppendLine(l)
        Next
        Return sb.ToString().TrimEnd()
    End Function

    ''' <summary>
    ''' Analyzes formatting of the specified range and returns SetBubbles-format string
    ''' (text@@comment§§§text@@comment...) for all suspicious spans found.
    ''' </summary>
    ''' <param name="rng">Word range to analyze.</param>
    ''' <returns>SetBubbles format string with all findings; empty string if no suspicions.</returns>
    ''' <remarks>
    ''' Calls AnalyzeFormattingSuspicion to run all heuristic checks, then formats results
    ''' as text@@reason§§§... for consumption by SetBubbles. Snippets are extracted from
    ''' SAME story using relative offsets to avoid misalignment across story boundaries.
    ''' </remarks>
    Private Function BuildSuspicionBubbleString(ByVal rng As Microsoft.Office.Interop.Word.Range) As System.String
        Try
            Dim findings As List(Of SuspiciousSpan) = AnalyzeFormattingSuspicion(rng)
            If findings Is Nothing OrElse findings.Count = 0 Then Return String.Empty

            Dim sb As New System.Text.StringBuilder()

            For Each f In findings
                Dim snippet As String = f.Snippet
                If String.IsNullOrEmpty(snippet) AndAlso f.Length > 0 Then
                    ' Reconstruct snippet from SAME story using relative offsets
                    Try
                        snippet = SliceByRel(rng, f.StartIndex, f.Length).Text
                    Catch
                        snippet = String.Empty
                    End Try
                End If

                sb.Append(snippet).Append("@@").Append(f.Reason).Append("§§§")
            Next

            If sb.Length >= 3 AndAlso sb.ToString().EndsWith("§§§", StringComparison.Ordinal) Then
                sb.Length -= 3
            End If
            Return sb.ToString()
        Catch
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Performs comprehensive formatting-based heuristic analysis to detect hidden/obfuscated text.
    ''' Runs 11 separate checks with progress tracking and cancellation support.
    ''' </summary>
    ''' <param name="rng">Word range to analyze.</param>
    ''' <returns>Deduplicated list of suspicious spans found.</returns>
    ''' <remarks>
    ''' Checks performed (all aggregate contiguous runs):
    '''  1. Word Hidden text (Font.Hidden = True/wdUndefined) - revealed in red
    '''  2. Very small font size (< 3pt)
    '''  3. Font color matching paragraph shading color
    '''  4. White-on-white text (explicit or Auto-resolved near-white)
    '''  5. Font color matching highlight color
    '''  6. Extreme font scaling (<= 10%)
    '''  7. Negative character spacing (< -2pt)
    '''  8. Font color matching table cell shading
    '''  9. White-on-white in table cells
    ''' 10. Zero-width and Bidi control characters
    ''' 11. Field codes with formatting switches (MERGEFORMAT, CHARFORMAT)
    ''' Progress is tracked per-character for hidden text detection and per-word for other checks.
    ''' Operation can be cancelled via ProgressBarModule.CancelOperation at any time.
    ''' Minimum finding length is 3 characters (configurable via MinFindingLen constant).
    ''' </remarks>
    Private Function AnalyzeFormattingSuspicion(ByVal rng As Microsoft.Office.Interop.Word.Range) As System.Collections.Generic.List(Of SuspiciousSpan)
        Dim findings As New System.Collections.Generic.List(Of SuspiciousSpan)()

        ' Initialize progress/cancel
        Try
            ProgressBarModule.CancelOperation = False
        Catch
            ' Ignore if progress module unavailable
        End Try

        Dim baseStart As Integer = rng.Start
        Const TinyFontPt As Single = 3.0F
        Const MinFindingLen As Integer = 3

        ' Calculate total work units for progress bar
        Dim charCount As Integer = 0
        Dim wordCount As Integer = 0
        Try : charCount = If(rng IsNot Nothing AndAlso rng.Characters IsNot Nothing, rng.Characters.Count, 0) : Catch : charCount = 0 : End Try
        Try : wordCount = If(rng IsNot Nothing AndAlso rng.Words IsNot Nothing, rng.Words.Count, 0) : Catch : wordCount = 0 : End Try

        Const HeuristicPasses As Integer = 10 ' Number of AddRuns calls below
        Dim totalUnits As Integer = charCount + (wordCount * HeuristicPasses)
        If totalUnits <= 0 Then totalUnits = 1

        ' Start progress window
        Try
            ProgressBarModule.GlobalProgressMax = totalUnits
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = "Scanning hidden text…"
            ProgressBarModule.ShowProgressBarInSeparateThread("Analyzing hidden/obfuscated text", "Preparing…")
        Catch
            ' Ignore UI issues
        End Try

        ' Check 1: Word Hidden text (Font.Hidden) - only if any hidden formatting exists
        If ProgressBarModule.CancelOperation Then Return findings
        Dim hiddenState As Integer = 0
        Try : hiddenState = CInt(rng.Font.Hidden) : Catch : hiddenState = 0 : End Try
        If hiddenState <> 0 Then
            AddHiddenRunsByFind(rng, baseStart, MinFindingLen, findings)
        End If

        ' Helper lambda: determines if range contains visible (non-whitespace) text
        Dim isVisibleToken As Func(Of Microsoft.Office.Interop.Word.Range, Boolean) =
            Function(w As Microsoft.Office.Interop.Word.Range)
                Dim t = w.Text
                Return Not String.IsNullOrEmpty(t) AndAlso t.Trim().Length > 0
            End Function

        ' Check 2: Very small font size
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking very small font size…"
        AddRuns(rng,
            Function(w)
                If Not isVisibleToken(w) Then Return False
                Return SafePt(w.Font.Size, 11.0F) < TinyFontPt
            End Function,
            "Very small font size",
            baseStart, MinFindingLen, findings)

        ' Check 3: Font color matching paragraph shading
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking font vs paragraph shading…"
        AddRuns(rng,
            Function(w)
                If Not isVisibleToken(w) Then Return False
                Return HasMeaningfulShading(w) AndAlso FontEqualsShadingColorIndex(w)
            End Function,
            "Font color equals background shading color",
            baseStart, MinFindingLen, findings)

        ' Check 4: White-on-white text
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking white-on-white…"
        AddRuns(rng,
            Function(w)
                If Not isVisibleToken(w) Then Return False
                Return IsWhiteOnWhite(w)
            End Function,
            "Likely white-on-white (near-invisible) text",
            baseStart, MinFindingLen, findings)

        ' Check 5: Font color matching highlight
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking font vs highlight color…"
        AddRuns(rng,
            Function(w)
                If Not isVisibleToken(w) Then Return False
                Return HasMeaningfulHighlight(w) AndAlso FontEqualsHighlightColorIndex(w)
            End Function,
            "Font color equals highlight color (camouflage)",
            baseStart, MinFindingLen, findings)

        ' Check 6: Extreme font scaling
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking extreme font scaling…"
        AddRuns(rng,
            Function(w)
                If Not isVisibleToken(w) Then Return False
                Return SafePercent(w.Font.Scaling, 100) <= 10
            End Function,
            "Extreme font scaling (condensed)",
            baseStart, MinFindingLen, findings)

        ' Check 7: Negative character spacing
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking negative character spacing…"
        AddRuns(rng,
            Function(w)
                If Not isVisibleToken(w) Then Return False
                Return SafePercent(w.Font.Spacing, 0) < -2
            End Function,
            "Very negative character spacing",
            baseStart, MinFindingLen, findings)

        ' Check 8: Font color matching table cell shading
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking font vs table cell shading…"
        AddRuns(rng,
            Function(w)
                Dim t = w.Text : If String.IsNullOrEmpty(t) OrElse t.Trim().Length = 0 Then Return False
                Return HasMeaningfulCellShading(w) AndAlso FontEqualsCellShadingColorIndex(w)
            End Function,
            "Font color equals table cell background color",
            baseStart, MinFindingLen, findings)

        ' Check 9: White-on-white in table cells
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking white-on-white in table cells…"
        AddRuns(rng,
            Function(w)
                Dim t = w.Text : If String.IsNullOrEmpty(t) OrElse t.Trim().Length = 0 Then Return False
                If Not HasMeaningfulCellShading(w) Then Return False
                Dim fontIsWhite As Boolean =
                    (w.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite) OrElse
                    IsLikelyInvisibleOnWhite(SafeWdColorToRgb(w.Font.Color))
                Return fontIsWhite AndAlso w.Cells(1).Shading.BackgroundPatternColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite
            End Function,
            "Likely white-on-white text in table cell",
            baseStart, MinFindingLen, findings)

        ' Check 10: Zero-width and Bidi control characters
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking zero-width/Bidi controls…"
        AddRuns(rng,
            Function(w)
                Dim raw As String = w.Text
                If String.IsNullOrEmpty(raw) Then Return False
                Return ContainsZeroWidthOrBidi(raw) AndAlso raw.Trim().Length > 0
            End Function,
            "Zero-width/Bidi control characters present",
            baseStart, MinFindingLen, findings)

        ' Check 11: Field codes with formatting switches
        If ProgressBarModule.CancelOperation Then Return findings
        ProgressBarModule.GlobalProgressLabel = "Checking field code formatting switches…"
        AddRuns(rng,
            Function(w)
                Try
                    If w.Fields Is Nothing OrElse w.Fields.Count = 0 Then Return False
                    For Each f As Microsoft.Office.Interop.Word.Field In w.Fields
                        If f Is Nothing OrElse f.Code Is Nothing OrElse f.Code.Text Is Nothing Then Continue For
                        Dim code As String = f.Code.Text
                        If code.IndexOf("\* MERGEFORMAT", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           code.IndexOf("\* CHARFORMAT", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Return True
                        End If
                    Next
                    Return False
                Catch
                    Return False
                End Try
            End Function,
            "Field code with formatting switches (may hide text)",
            baseStart, MinFindingLen, findings)

        Dim result = DeduplicateFindings(findings)

        ' Complete progress and close window
        Try
            ProgressBarModule.GlobalProgressLabel = "Completed"
            ProgressBarModule.GlobalProgressValue = ProgressBarModule.GlobalProgressMax
            System.Threading.Tasks.Task.Delay(500)
            ProgressBarModule.CancelOperation = True
        Catch
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Creates a Range slice within the SAME story as rng using absolute story coordinates.
    ''' </summary>
    ''' <param name="rng">Source range defining the story context.</param>
    ''' <param name="absStart">Absolute start position within story.</param>
    ''' <param name="absEnd">Absolute end position within story.</param>
    ''' <returns>Duplicate range with specified absolute start/end positions.</returns>
    ''' <remarks>
    ''' Essential for multi-story documents to avoid cross-story position misalignment.
    ''' Always operates within the same story type as source range.
    ''' </remarks>
    Private Function SliceInSameStory(rng As Word.Range, absStart As Integer, absEnd As Integer) As Word.Range
        Dim slice As Word.Range = rng.Duplicate
        slice.Start = absStart
        slice.End = absEnd
        Return slice
    End Function

    ''' <summary>
    ''' Creates a Range slice using relative offsets from rng.Start within the SAME story.
    ''' </summary>
    ''' <param name="rng">Source range defining story context and base position.</param>
    ''' <param name="relStart">Relative start offset from rng.Start.</param>
    ''' <param name="length">Length of slice in characters.</param>
    ''' <returns>Duplicate range with calculated absolute positions.</returns>
    ''' <remarks>
    ''' Converts relative offsets (used in SuspiciousSpan) to absolute story positions
    ''' for correct slicing. Negative offsets are clamped to zero.
    ''' </remarks>
    Private Function SliceByRel(rng As Word.Range, relStart As Integer, length As Integer) As Word.Range
        Dim absStart As Integer = rng.Start + System.Math.Max(0, relStart)
        Dim absEnd As Integer = absStart + System.Math.Max(0, length)
        Return SliceInSameStory(rng, absStart, absEnd)
    End Function

    ''' <summary>
    ''' Detects Word Hidden text by character-by-character scanning, merges contiguous runs,
    ''' reveals them (unhides and colors red), and records findings.
    ''' </summary>
    ''' <param name="rng">Word range to scan for hidden text.</param>
    ''' <param name="baseStart">Absolute start position of parent range (for relative offset calculation).</param>
    ''' <param name="minLen">Minimum run length (in characters) to report.</param>
    ''' <param name="findings">List to append findings to.</param>
    ''' <remarks>
    ''' Fast pre-check: only proceeds if rng.Font.Hidden indicates any hidden formatting exists.
    ''' Progress: increments per character (updates label every 500 characters).
    ''' Cancellation: honors ProgressBarModule.CancelOperation at every character.
    ''' Contiguity: merges adjacent hidden characters into single run; splits on gaps.
    ''' Reveal: calls RevealHiddenRun (unhide + red color) for each qualifying run before recording.
    ''' </remarks>
    Private Sub AddHiddenRunsByFind(
        ByVal rng As Microsoft.Office.Interop.Word.Range,
        ByVal baseStart As Integer,
        ByVal minLen As Integer,
        ByVal findings As System.Collections.Generic.List(Of SuspiciousSpan)
    )
        ' Fast skip: only proceed if range has any hidden formatting
        Dim hiddenState As Integer = 0
        Try : hiddenState = CInt(rng.Font.Hidden) : Catch : hiddenState = 0 : End Try
        If hiddenState = 0 Then Exit Sub

        Try
            Dim chars = rng.Characters
            If chars Is Nothing OrElse chars.Count = 0 Then Exit Sub

            Dim inRun As Boolean = False
            Dim runStartAbs As Integer = -1
            Dim lastEndAbs As Integer = -1
            Dim stepCounter As Integer = 0

            For i As Integer = 1 To chars.Count
                If ProgressBarModule.CancelOperation Then Exit Sub

                Dim ch As Microsoft.Office.Interop.Word.Range = Nothing
                Dim isHidden As Boolean = False
                Try
                    ch = chars(i)
                    Dim hiddenVal As Integer = 0
                    Try
                        hiddenVal = CInt(ch.Font.Hidden)
                    Catch
                        hiddenVal = 0
                    End Try
                    isHidden = (hiddenVal <> 0)
                Catch
                    isHidden = False
                End Try

                If isHidden Then
                    If Not inRun Then
                        ' Start new run
                        inRun = True
                        runStartAbs = ch.Start
                        lastEndAbs = ch.End
                    Else
                        ' Continue run if contiguous, otherwise flush previous
                        If ch.Start = lastEndAbs Then
                            lastEndAbs = ch.End
                        Else
                            Dim runLen As Integer = System.Math.Max(0, lastEndAbs - runStartAbs)
                            If runLen >= minLen Then
                                RevealHiddenRun(rng, runStartAbs, lastEndAbs)
                                AddRunFinding(findings, "Hidden text span (revealed in red)", runStartAbs, lastEndAbs, baseStart, rng)
                            End If
                            runStartAbs = ch.Start
                            lastEndAbs = ch.End
                        End If
                    End If
                ElseIf inRun Then
                    ' End of run - flush if meets minimum length
                    Dim runLen As Integer = System.Math.Max(0, lastEndAbs - runStartAbs)
                    If runLen >= minLen Then
                        RevealHiddenRun(rng, runStartAbs, lastEndAbs)
                        AddRunFinding(findings, "Hidden text span (revealed in red)", runStartAbs, lastEndAbs, baseStart, rng)
                    End If
                    inRun = False
                    runStartAbs = -1
                    lastEndAbs = -1
                End If

                ' Update progress per character
                stepCounter += 1
                Try
                    ProgressBarModule.GlobalProgressValue = System.Math.Min(ProgressBarModule.GlobalProgressValue + 1, ProgressBarModule.GlobalProgressMax)
                    If (stepCounter Mod 500) = 0 Then
                        ProgressBarModule.GlobalProgressLabel = $"Scanning hidden text… ({stepCounter}/{System.Math.Max(1, chars.Count)})"
                    End If
                Catch
                End Try
            Next

            ' Flush trailing run
            If inRun AndAlso runStartAbs >= 0 AndAlso lastEndAbs > runStartAbs Then
                Dim runLen As Integer = System.Math.Max(0, lastEndAbs - runStartAbs)
                If runLen >= minLen Then
                    RevealHiddenRun(rng, runStartAbs, lastEndAbs)
                    AddRunFinding(findings, "Hidden text span (revealed in red)", runStartAbs, lastEndAbs, baseStart, rng)
                End If
            End If
        Catch
            ' Best effort - ignore errors
        End Try
    End Sub

    ''' <summary>
    ''' Unhides and colors red a hidden text span without altering text content or positions.
    ''' </summary>
    ''' <param name="selectionRange">Parent range defining story context.</param>
    ''' <param name="startAbs">Absolute start position of hidden span.</param>
    ''' <param name="endAbs">Absolute end position of hidden span.</param>
    ''' <remarks>
    ''' Uses SliceInSameStory to ensure correct positioning within story.
    ''' Sets Font.Hidden = 0 (False) and Font.ColorIndex = wdRed.
    ''' Formatting failures are silently ignored to allow batch processing to continue.
    ''' </remarks>
    Private Sub RevealHiddenRun(ByVal selectionRange As Microsoft.Office.Interop.Word.Range,
                                ByVal startAbs As Integer,
                                ByVal endAbs As Integer)
        Try
            Dim dr As Microsoft.Office.Interop.Word.Range = SliceInSameStory(selectionRange, startAbs, endAbs)
            dr.Font.Hidden = 0
            dr.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdRed
        Catch
            ' Ignore formatting failures
        End Try
    End Sub

    ''' <summary>
    ''' Aggregates contiguous Word ranges satisfying predicate into single findings per run.
    ''' Iterates words, merges adjacent matching words, emits findings >= minLen.
    ''' </summary>
    ''' <param name="rng">Word range to analyze.</param>
    ''' <param name="predicate">Function returning True for suspicious words.</param>
    ''' <param name="reason">Human-readable reason string for findings.</param>
    ''' <param name="baseStart">Absolute start position of parent range.</param>
    ''' <param name="minLen">Minimum run length (characters) to report.</param>
    ''' <param name="findings">List to append findings to.</param>
    ''' <remarks>
    ''' Progress: increments GlobalProgressValue per word; supports cancellation.
    ''' Contiguity: determined by adjacent word Start/End positions.
    ''' Trailing runs are flushed after loop completes.
    ''' Predicate exceptions are caught and treated as False (skip word).
    ''' </remarks>
    Private Sub AddRuns(
        ByVal rng As Microsoft.Office.Interop.Word.Range,
        ByVal predicate As Func(Of Microsoft.Office.Interop.Word.Range, Boolean),
        ByVal reason As String,
        ByVal baseStart As Integer,
        ByVal minLen As Integer,
        ByVal findings As System.Collections.Generic.List(Of SuspiciousSpan)
    )
        Dim words = rng.Words
        If words Is Nothing OrElse words.Count = 0 Then Exit Sub

        Dim inRun As Boolean = False
        Dim runStartAbs As Integer = -1
        Dim runEndAbs As Integer = -1

        For i As Integer = 1 To words.Count
            If ProgressBarModule.CancelOperation Then Exit Sub

            Dim w As Microsoft.Office.Interop.Word.Range = words(i)
            Dim ok As Boolean = False
            Try
                ok = predicate(w)
            Catch
                ok = False
            End Try

            If ok Then
                If Not inRun Then
                    inRun = True
                    runStartAbs = w.Start
                End If
                runEndAbs = w.End
            ElseIf inRun Then
                Dim runLen As Integer = System.Math.Max(0, runEndAbs - runStartAbs)
                If runLen >= minLen Then
                    AddRunFinding(findings, reason, runStartAbs, runEndAbs, baseStart, rng)
                End If
                inRun = False
                runStartAbs = -1
                runEndAbs = -1
            End If

            ' Update progress per word
            Try
                ProgressBarModule.GlobalProgressValue = System.Math.Min(ProgressBarModule.GlobalProgressValue + 1, ProgressBarModule.GlobalProgressMax)
            Catch
            End Try
        Next

        ' Flush trailing run
        If inRun Then
            Dim runLen As Integer = System.Math.Max(0, runEndAbs - runStartAbs)
            If runLen >= minLen Then
                AddRunFinding(findings, reason, runStartAbs, runEndAbs, baseStart, rng)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Detects white-on-white text considering both ColorIndex and resolved RGB values.
    ''' Checks against default background, highlight, and paragraph shading.
    ''' </summary>
    ''' <param name="w">Word range to check.</param>
    ''' <returns>True if text is effectively invisible (white-on-white); otherwise False.</returns>
    ''' <remarks>
    ''' Handles explicit wdWhite ColorIndex and Auto/ByAuthor colors resolved to near-white RGB.
    ''' Checks: (1) no shading/highlight, (2) white highlight, (3) white shading.
    ''' RGB threshold: average >= 245 (see IsLikelyInvisibleOnWhite).
    ''' </remarks>
    Private Function IsWhiteOnWhite(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            Dim fci = w.Font.ColorIndex
            Dim fontIsWhiteIdx As Boolean = (fci = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite)

            ' Resolve Auto/ByAuthor to RGB for near-white detection
            Dim rgb As Integer = SafeWdColorToRgb(w.Font.Color)
            Dim nearWhite As Boolean = IsLikelyInvisibleOnWhite(rgb)

            Dim hasShade As Boolean = HasMeaningfulShading(w)
            Dim hasHl As Boolean = HasMeaningfulHighlight(w)

            ' Check explicit white ColorIndex
            If fontIsWhiteIdx Then
                If Not hasShade AndAlso Not hasHl Then Return True
                If hasHl AndAlso w.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite Then Return True
                If hasShade AndAlso w.Shading.BackgroundPatternColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite Then Return True
            End If

            ' Check Auto/ByAuthor resolved to near-white
            If (fci = Microsoft.Office.Interop.Word.WdColorIndex.wdAuto OrElse
                fci = Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor) AndAlso nearWhite Then
                If Not hasShade AndAlso Not hasHl Then Return True
                If hasHl AndAlso w.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite Then Return True
                If hasShade AndAlso w.Shading.BackgroundPatternColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite Then Return True
            End If

            Return False
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determines whether range has meaningful table cell shading (not Auto/ByAuthor/NoHighlight).
    ''' </summary>
    ''' <param name="w">Word range to check (must be inside table cell).</param>
    ''' <returns>True if cell has explicit background shading; otherwise False.</returns>
    ''' <remarks>
    ''' Returns False if range not in table or Cells collection unavailable.
    ''' Used to detect font camouflage against table cell backgrounds.
    ''' </remarks>
    Private Function HasMeaningfulCellShading(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            If w.Cells Is Nothing OrElse w.Cells.Count = 0 Then Return False
            Dim sh = w.Cells(1).Shading
            Return sh IsNot Nothing AndAlso
                   sh.BackgroundPatternColorIndex <> Microsoft.Office.Interop.Word.WdColorIndex.wdAuto AndAlso
                   sh.BackgroundPatternColorIndex <> Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor AndAlso
                   sh.BackgroundPatternColorIndex <> Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determines whether font ColorIndex matches table cell shading ColorIndex (camouflage).
    ''' </summary>
    ''' <param name="w">Word range to check (must be inside table cell).</param>
    ''' <returns>True if font and cell background colors match; otherwise False.</returns>
    ''' <remarks>
    ''' Returns False if not in table, or font color is Auto/ByAuthor (unresolved).
    ''' Compares ColorIndex values directly (integer cast for safety).
    ''' </remarks>
    Private Function FontEqualsCellShadingColorIndex(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            If w.Cells Is Nothing OrElse w.Cells.Count = 0 Then Return False
            Dim fci = w.Font.ColorIndex
            If fci = Microsoft.Office.Interop.Word.WdColorIndex.wdAuto OrElse
               fci = Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor Then
                Return False
            End If
            Return CInt(fci) = CInt(w.Cells(1).Shading.BackgroundPatternColorIndex)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Creates a SuspiciousSpan finding and adds it to the findings list.
    ''' Calculates relative offset, extracts snippet from same-story slice.
    ''' </summary>
    ''' <param name="findings">List to append finding to.</param>
    ''' <param name="reason">Human-readable reason string.</param>
    ''' <param name="runStartAbs">Absolute start position in story.</param>
    ''' <param name="runEndAbs">Absolute end position in story.</param>
    ''' <param name="baseStart">Absolute start of parent range (for relative calculation).</param>
    ''' <param name="selectionRange">Parent range defining story context.</param>
    ''' <remarks>
    ''' Primary overload used by most heuristic checks.
    ''' Snippet extracted via SliceInSameStory to ensure correct story positioning.
    ''' </remarks>
    Private Sub AddRunFinding(
        ByVal findings As System.Collections.Generic.List(Of SuspiciousSpan),
        ByVal reason As String,
        ByVal runStartAbs As Integer,
        ByVal runEndAbs As Integer,
        ByVal baseStart As Integer,
        ByVal selectionRange As Microsoft.Office.Interop.Word.Range
    )
        Dim relStart As Integer = System.Math.Max(0, runStartAbs - baseStart)
        Dim length As Integer = System.Math.Max(0, runEndAbs - runStartAbs)
        Dim snippet As String = SliceInSameStory(selectionRange, runStartAbs, runEndAbs).Text

        findings.Add(New SuspiciousSpan With {
            .Reason = reason,
            .StartIndex = relStart,
            .Length = length,
            .Snippet = snippet
        })
    End Sub

    ''' <summary>
    ''' Creates a SuspiciousSpan finding with fallback to substring if Range slice fails.
    ''' Alternative overload accepting fullText for backward compatibility.
    ''' </summary>
    ''' <param name="findings">List to append finding to.</param>
    ''' <param name="reason">Human-readable reason string.</param>
    ''' <param name="runStartAbs">Absolute start position in story.</param>
    ''' <param name="runEndAbs">Absolute end position in story.</param>
    ''' <param name="baseStart">Absolute start of parent range.</param>
    ''' <param name="fullText">Fallback full text string for substring extraction.</param>
    ''' <param name="selectionRange">Parent range defining story context.</param>
    ''' <remarks>
    ''' Tries SliceInSameStory first; falls back to substring if slice fails.
    ''' Substring fallback may produce incorrect results with hidden/control characters.
    ''' </remarks>
    Private Sub AddRunFinding(
        ByVal findings As System.Collections.Generic.List(Of SuspiciousSpan),
        ByVal reason As String,
        ByVal runStartAbs As Integer,
        ByVal runEndAbs As Integer,
        ByVal baseStart As Integer,
        ByVal fullText As String,
        ByVal selectionRange As Microsoft.Office.Interop.Word.Range
    )
        Dim relStart As Integer = System.Math.Max(0, runStartAbs - baseStart)
        Dim length As Integer = System.Math.Max(0, runEndAbs - runStartAbs)

        Dim snippet As String
        Try
            snippet = SliceInSameStory(selectionRange, runStartAbs, runEndAbs).Text
        Catch
            ' Fallback to substring if slice fails
            If length > 0 AndAlso relStart + length <= fullText.Length Then
                snippet = fullText.Substring(relStart, length)
            Else
                snippet = String.Empty
            End If
        End Try

        findings.Add(New SuspiciousSpan With {
            .Reason = reason,
            .StartIndex = relStart,
            .Length = length,
            .Snippet = snippet
        })
    End Sub

    ''' <summary>
    ''' Determines whether range has meaningful paragraph shading (not Auto/None/ByAuthor).
    ''' </summary>
    ''' <param name="w">Word range to check.</param>
    ''' <returns>True if paragraph has explicit background shading; otherwise False.</returns>
    ''' <remarks>
    ''' Checks Shading.Texture != wdTextureNone and BackgroundPatternColorIndex is explicit.
    ''' Used to detect font camouflage against paragraph backgrounds.
    ''' </remarks>
    Private Function HasMeaningfulShading(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            Return (w.Shading IsNot Nothing) AndAlso
                   (w.Shading.Texture <> Microsoft.Office.Interop.Word.WdTextureIndex.wdTextureNone) AndAlso
                   (w.Shading.BackgroundPatternColorIndex <> Microsoft.Office.Interop.Word.WdColorIndex.wdAuto) AndAlso
                   (w.Shading.BackgroundPatternColorIndex <> Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor) AndAlso
                   (w.Shading.BackgroundPatternColorIndex <> Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determines whether font ColorIndex matches paragraph shading ColorIndex (camouflage).
    ''' </summary>
    ''' <param name="w">Word range to check.</param>
    ''' <returns>True if font and background colors match; otherwise False.</returns>
    ''' <remarks>
    ''' Returns False if font color is Auto/ByAuthor (unresolved).
    ''' Compares ColorIndex values directly (integer cast for safety).
    ''' </remarks>
    Private Function FontEqualsShadingColorIndex(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            Dim fci = w.Font.ColorIndex
            If fci = Microsoft.Office.Interop.Word.WdColorIndex.wdAuto OrElse
               fci = Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor Then
                Return False
            End If
            Return CInt(fci) = CInt(w.Shading.BackgroundPatternColorIndex)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determines whether range has meaningful highlight (not NoHighlight or ByAuthor).
    ''' </summary>
    ''' <param name="w">Word range to check.</param>
    ''' <returns>True if range has explicit highlight color; otherwise False.</returns>
    ''' <remarks>
    ''' Used to detect font camouflage against highlight backgrounds.
    ''' </remarks>
    Private Function HasMeaningfulHighlight(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            Dim hi = w.HighlightColorIndex
            Return hi <> Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight AndAlso
                   hi <> Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Determines whether font ColorIndex matches highlight ColorIndex (camouflage).
    ''' </summary>
    ''' <param name="w">Word range to check.</param>
    ''' <returns>True if font and highlight colors match; otherwise False.</returns>
    ''' <remarks>
    ''' Returns False if highlight is NoHighlight/ByAuthor or font is Auto/ByAuthor.
    ''' Compares ColorIndex values directly (integer cast for safety).
    ''' </remarks>
    Private Function FontEqualsHighlightColorIndex(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            Dim hi = w.HighlightColorIndex
            Dim fci = w.Font.ColorIndex
            If hi = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight OrElse
               hi = Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor OrElse
               fci = Microsoft.Office.Interop.Word.WdColorIndex.wdAuto OrElse
               fci = Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor Then
                Return False
            End If
            Return CInt(hi) = CInt(fci)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Overload of AddRuns accepting fullText parameter (legacy/compatibility signature).
    ''' NOT USED - duplicate signature retained for backward compatibility only.
    ''' </summary>
    ''' <remarks>
    ''' This overload is not actively called by current code but remains to avoid breaking
    ''' any potential external references. Consider removing in future refactoring.
    ''' </remarks>
    Private Sub AddRuns(
        ByVal rng As Microsoft.Office.Interop.Word.Range,
        ByVal predicate As Func(Of Microsoft.Office.Interop.Word.Range, Boolean),
        ByVal reason As String,
        ByVal baseStart As Integer,
        ByVal fullText As String,
        ByVal minLen As Integer,
        ByVal findings As System.Collections.Generic.List(Of SuspiciousSpan)
    )
        Dim words = rng.Words
        If words Is Nothing OrElse words.Count = 0 Then Exit Sub

        Dim inRun As Boolean = False
        Dim runStartAbs As Integer = -1
        Dim runEndAbs As Integer = -1

        For i As Integer = 1 To words.Count
            Dim w As Microsoft.Office.Interop.Word.Range = words(i)
            Dim ok As Boolean = False
            Try
                ok = predicate(w)
            Catch
                ok = False
            End Try

            If ok Then
                If Not inRun Then
                    inRun = True
                    runStartAbs = w.Start
                End If
                runEndAbs = w.End
            ElseIf inRun Then
                Dim runLen As Integer = System.Math.Max(0, runEndAbs - runStartAbs)
                If runLen >= minLen Then
                    AddRunFinding(findings, reason, runStartAbs, runEndAbs, baseStart, fullText, rng)
                End If
                inRun = False
                runStartAbs = -1
                runEndAbs = -1
            End If
        Next

        If inRun Then
            Dim runLen As Integer = System.Math.Max(0, runEndAbs - runStartAbs)
            If runLen >= minLen Then
                AddRunFinding(findings, reason, runStartAbs, runEndAbs, baseStart, fullText, rng)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Determines whether RGB color value is near-white (effectively invisible on white background).
    ''' </summary>
    ''' <param name="rgb">RGB color value (R in lowest byte, G in middle, B in highest).</param>
    ''' <returns>True if average RGB >= 245; otherwise False.</returns>
    ''' <remarks>
    ''' Threshold of 245 chosen empirically to catch near-white Auto colors while avoiding false positives.
    ''' Returns False for invalid RGB (-1 from SafeWdColorToRgb conversion failures).
    ''' </remarks>
    Private Function IsLikelyInvisibleOnWhite(ByVal rgb As System.Int32) As System.Boolean
        If rgb = -1 Then Return False
        Dim r As System.Int32 = (rgb And &HFF)
        Dim g As System.Int32 = ((rgb >> 8) And &HFF)
        Dim b As System.Int32 = ((rgb >> 16) And &HFF)
        Dim avg As System.Int32 = (r + g + b) \ 3
        Return avg >= 245
    End Function

    ''' <summary>
    ''' Converts Word WdColor value (BGR format) to standard RGB integer.
    ''' </summary>
    ''' <param name="wdColor">Word color value (Object type to handle variants).</param>
    ''' <returns>RGB integer (R in lowest byte); -1 if conversion fails.</returns>
    ''' <remarks>
    ''' Word stores colors in BGR format (Blue in lowest byte).
    ''' Performs byte swapping: BGR -> RGB for standard processing.
    ''' Returns -1 on conversion errors (used as sentinel by callers).
    ''' </remarks>
    Private Function SafeWdColorToRgb(ByVal wdColor As System.Object) As System.Int32
        Try
            Dim bgr As System.Int32 = System.Convert.ToInt32(wdColor, Globalization.CultureInfo.InvariantCulture)
            Dim rr As System.Int32 = (bgr And &HFF)
            Dim gg As System.Int32 = ((bgr >> 8) And &HFF)
            Dim bb As System.Int32 = ((bgr >> 16) And &HFF)
            Dim rgb As System.Int32 = (rr << 16) Or (gg << 8) Or bb
            Return rgb
        Catch
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Safely converts font size value to Single with fallback.
    ''' </summary>
    ''' <param name="size">Font size value (may be Object/Variant from Word API).</param>
    ''' <param name="fallbackPt">Fallback value in points if conversion fails.</param>
    ''' <returns>Font size as Single; fallbackPt if conversion fails.</returns>
    ''' <remarks>
    ''' Handles Word's variant font size values that may be undefined or mixed.
    ''' Uses InvariantCulture to avoid locale-specific parsing issues.
    ''' </remarks>
    Private Function SafePt(ByVal size As System.Object, ByVal fallbackPt As System.Single) As System.Single
        Try
            Return System.Convert.ToSingle(size, Globalization.CultureInfo.InvariantCulture)
        Catch
            Return fallbackPt
        End Try
    End Function

    ''' <summary>
    ''' Safely converts percentage/spacing value to Integer with fallback.
    ''' </summary>
    ''' <param name="val">Percentage or spacing value (may be Object/Variant from Word API).</param>
    ''' <param name="fallback">Fallback value if conversion fails.</param>
    ''' <returns>Value as Integer; fallback if conversion fails.</returns>
    ''' <remarks>
    ''' Used for Font.Scaling and Font.Spacing which may be undefined or mixed in selections.
    ''' Uses InvariantCulture to avoid locale-specific parsing issues.
    ''' </remarks>
    Private Function SafePercent(ByVal val As System.Object, ByVal fallback As System.Int32) As System.Int32
        Try
            Return System.Convert.ToInt32(val, Globalization.CultureInfo.InvariantCulture)
        Catch
            Return fallback
        End Try
    End Function

    ''' <summary>
    ''' Determines whether string contains zero-width or bidirectional control characters.
    ''' </summary>
    ''' <param name="s">String to check for invisible characters.</param>
    ''' <returns>True if string contains zero-width or Bidi controls; otherwise False.</returns>
    ''' <remarks>
    ''' Detects the following Unicode characters:
    ''' Zero-width: U+200B (ZWSP), U+200C (ZWNJ), U+200D (ZWJ)
    ''' Bidi isolates: U+2066 (LRI), U+2067 (RLI), U+2068 (FSI), U+2069 (PDI)
    ''' Bidi embeddings: U+202A (LRE), U+202B (RLE), U+202D (LRO), U+202E (RLO), U+202C (PDF)
    ''' These characters can be used for prompt injection or text manipulation attacks.
    ''' </remarks>
    Private Function ContainsZeroWidthOrBidi(ByVal s As System.String) As System.Boolean
        If s Is Nothing OrElse s.Length = 0 Then Return False
        For Each ch As System.Char In s
            Dim code As System.Int32 = System.Convert.ToInt32(ch)
            If code = &H200B OrElse code = &H200C OrElse code = &H200D OrElse
               code = &H2066 OrElse code = &H2067 OrElse code = &H2068 OrElse code = &H2069 OrElse
               code = &H202A OrElse code = &H202B OrElse code = &H202D OrElse code = &H202E OrElse code = &H202C Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Removes duplicate findings based on composite key (reason, position, length, snippet).
    ''' </summary>
    ''' <param name="input">List of findings potentially containing duplicates.</param>
    ''' <returns>Deduplicated list preserving first occurrence of each unique finding.</returns>
    ''' <remarks>
    ''' Duplicate key format: "Reason|StartIndex|Length|Snippet"
    ''' Uses StringComparer.Ordinal for case-sensitive, culture-invariant comparison.
    ''' Necessary because multiple heuristics may detect the same span (e.g., hidden + tiny font).
    ''' </remarks>
    Private Function DeduplicateFindings(ByVal input As System.Collections.Generic.List(Of SuspiciousSpan)) As System.Collections.Generic.List(Of SuspiciousSpan)
        Dim seen As System.Collections.Generic.HashSet(Of System.String) = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        Dim result As System.Collections.Generic.List(Of SuspiciousSpan) = New System.Collections.Generic.List(Of SuspiciousSpan)()
        For Each f As SuspiciousSpan In input
            Dim key As System.String = f.Reason & "|" & f.StartIndex.ToString(Globalization.CultureInfo.InvariantCulture) & "|" & f.Length.ToString(Globalization.CultureInfo.InvariantCulture) & "|" & f.Snippet
            If Not seen.Contains(key) Then
                seen.Add(key)
                result.Add(f)
            End If
        Next
        Return result
    End Function

    ''' <summary>
    ''' Internal data structure representing a suspicious text span detected by heuristics.
    ''' </summary>
    ''' <remarks>
    ''' Used to accumulate findings from AnalyzeFormattingSuspicion before conversion to
    ''' SetBubbles format (text@@reason§§§...) for comment insertion.
    ''' StartIndex is relative to parent range start (not absolute document position).
    ''' Snippet may be empty if extraction fails; BuildSuspicionBubbleString reconstructs if needed.
    ''' </remarks>
    Private NotInheritable Class SuspiciousSpan
        ''' <summary>Human-readable reason describing the suspicion (e.g., "Hidden text span (revealed in red)").</summary>
        Public Property Reason As System.String

        ''' <summary>Relative start position within parent range (not absolute document position).</summary>
        Public Property StartIndex As System.Int32

        ''' <summary>Length of suspicious span in characters.</summary>
        Public Property Length As System.Int32

        ''' <summary>Text snippet extracted from span; may be empty if extraction fails.</summary>
        Public Property Snippet As System.String
    End Class

End Class
