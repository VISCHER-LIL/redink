' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Public Async Function FindHiddenPrompts() As System.Threading.Tasks.Task

        If INILoadFail() Then Return

        Dim Prefix As String = "-FHP"

        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection

            Dim JumpRoundA As Boolean = False
            Dim CheckAll As Boolean = False

            ' Cache the exact selection to reuse later (non-CheckAll path)
            Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.")
                Return
            End If

            If sel.Type = WdSelectionType.wdSelectionIP Then
                Dim answer As Integer = ShowCustomYesNoBox("You have not selected any text. Check the all text (including foot- and endnotes) or instead check a file?", "Check all text", "Check file")
                If answer = 0 Then Return
                If answer = 1 Then
                    app.Selection.WholeStory()
                    sel = app.Selection
                    CheckAll = True
                Else
                    DragDropFormLabel = "Document files (.txt, .docx, .pdf) or Powerpoint (.pptx)."
                    DragDropFormFilter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.pptx||" &
                                     "Text Files (*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                                     "Rich Text Files (*.rtf)|*.rtf|" &
                                     "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                                     "PDF Files (*.pdf)|*.pdf" &
                                     "Powerpoint Files (*.pptx)|*.pptx"

                    Dim FilePath As String = GetFileName()
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    If String.IsNullOrWhiteSpace(FilePath) Then
                        ShowCustomMessageBox("No file has been selected - will abort.")
                        Return
                    End If

                    Dim ext As String = IO.Path.GetExtension(FilePath).ToLowerInvariant()

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

            ' If user chose "Check all text" (no special file flow), iterate all stories.
            If CheckAll AndAlso Not JumpRoundA Then
                Dim originalStart As Integer = sel.Start
                Dim originalEnd As Integer = sel.End

                ' Build a flat list of all story ranges (including linked instances)
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

                ' Aggregate findings for non-commentable stories we must not use SetBubbles on.
                Dim footnoteLines As New List(Of String)()  ' For footnotes
                Dim endnoteLines As New List(Of String)()   ' For endnotes

                ' Pass 1: Deterministic checks (Round A) across ALL stories
                For Each s In storyList
                    Dim tr As Word.Range = s.rng
                    Dim st As Word.WdStoryType = s.st
                    If tr Is Nothing OrElse String.IsNullOrWhiteSpace(tr.Text) Then Continue For

                    Dim roundA As String = BuildSuspicionBubbleString(tr)
                    Debug.WriteLine($"FindHiddenPrompts[RoundA][{st}]: {roundA}")

                    If Not String.IsNullOrEmpty(roundA) Then
                        If IsFootnoteOrEndnote(st) Then
                            If st = Word.WdStoryType.wdFootnotesStory Then
                                footnoteLines.AddRange(ConvertSetBubblesToLines(roundA))
                            ElseIf st = Word.WdStoryType.wdEndnotesStory Then
                                endnoteLines.AddRange(ConvertSetBubblesToLines(roundA))
                            End If
                        ElseIf IsCommentableStory(st) Then
                            Try
                                tr.Select()
                                SetBubbles(roundA, app.Selection, True, Prefix)
                            Catch
                                ' Silently skip if comments fail for this story
                            End Try
                        End If
                    End If
                Next

                ' Pass 2: LLM checks (Round B) across ALL stories
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
                            If st = Word.WdStoryType.wdFootnotesStory Then
                                footnoteLines.AddRange(ConvertSetBubblesToLines(llmResult))
                            ElseIf st = Word.WdStoryType.wdEndnotesStory Then
                                endnoteLines.AddRange(ConvertSetBubblesToLines(llmResult))
                            End If
                        ElseIf IsCommentableStory(st) Then
                            Try
                                tr.Select()
                                SetBubbles(llmResult, app.Selection, True, Prefix)
                            Catch
                                ' Silently skip if comments fail for this story
                            End Try
                        End If
                    End If
                Next

                ' Add a single summary comment at the very end of the MAIN story for footnote/endnote findings.
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

            ' ===== Original behavior (selection-based OR JumpRoundA flow) =====

            Dim selStart As System.Int32 = sel.Start
            Dim selEnd As System.Int32 = sel.End
            Dim sameSel As Microsoft.Office.Interop.Word.Selection = Nothing

            If Not JumpRoundA Then
                Dim roundA As System.String = BuildSuspicionBubbleString(sel.Range)
                Debug.WriteLine("FindHiddenPrompts: Formatting-based findings: " & roundA)

                If Not System.String.IsNullOrEmpty(roundA) Then
                    ' Apply first-round bubbles
                    SetBubbles(roundA, sel, True, Prefix)
                End If

                ' Reselect the ORIGINAL span before (b)
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

            ' LLM check on the SAME selection as before
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

    Private Function IsFootnoteOrEndnote(st As Word.WdStoryType) As Boolean
        Return st = Word.WdStoryType.wdFootnotesStory OrElse st = Word.WdStoryType.wdEndnotesStory
    End Function

    Private Function IsCommentableStory(st As Word.WdStoryType) As Boolean
        ' Comments (and thus SetBubbles) are allowed in main body and usually text frames.
        ' They are not allowed in headers/footers/footnotes/endnotes.
        Return st = Word.WdStoryType.wdMainTextStory OrElse st = Word.WdStoryType.wdTextFrameStory
    End Function

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

    Private Function BuildNoticeText(scope As String, lines As List(Of String)) As String
        Dim sb As New System.Text.StringBuilder()
        ' Heading required: "Suspicious Footnote text:" or "Suspicious Endnote text:"
        sb.AppendLine($"Suspicious {scope} text:")
        For Each l In lines
            sb.AppendLine(l)
        Next
        Return sb.ToString().TrimEnd()
    End Function

    ' ===== (a) Build: text@@comment§§§text@@comment... =====
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


    ' ===== Formatting heuristics (includes Word Hidden text) =====

    Private Function AnalyzeFormattingSuspicion(ByVal rng As Microsoft.Office.Interop.Word.Range) As System.Collections.Generic.List(Of SuspiciousSpan)
        Dim findings As New System.Collections.Generic.List(Of SuspiciousSpan)()

        ' Initialize progress/cancel
        Try
            ProgressBarModule.CancelOperation = False
        Catch
            ' ignore
        End Try

        Dim baseStart As Integer = rng.Start
        Const TinyFontPt As Single = 3.0F
        Const MinFindingLen As Integer = 3 ' strictly > 2

        ' Pre-compute total work units for progress:
        Dim charCount As Integer = 0
        Dim wordCount As Integer = 0
        Try : charCount = If(rng IsNot Nothing AndAlso rng.Characters IsNot Nothing, rng.Characters.Count, 0) : Catch : charCount = 0 : End Try
        Try : wordCount = If(rng IsNot Nothing AndAlso rng.Words IsNot Nothing, rng.Words.Count, 0) : Catch : wordCount = 0 : End Try

        ' Heuristic passes that iterate words below (keep in sync with AddRuns calls)
        Const HeuristicPasses As Integer = 10
        Dim totalUnits As Integer = charCount + (wordCount * HeuristicPasses)
        If totalUnits <= 0 Then totalUnits = 1

        ' Start progress window
        Try
            ProgressBarModule.GlobalProgressMax = totalUnits
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressLabel = "Scanning hidden text…"
            ProgressBarModule.ShowProgressBarInSeparateThread("Analyzing hidden/obfuscated text", "Preparing…")
        Catch
            ' ignore UI issues
        End Try

        ' 1) Robust Hidden text spans (Ausgeblendet) — only if there is any hidden text (True or wdUndefined)
        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        Dim hiddenState As Integer = 0
        Try : hiddenState = CInt(rng.Font.Hidden) : Catch : hiddenState = 0 : End Try
        If hiddenState <> 0 Then
            AddHiddenRunsByFind(rng, baseStart, MinFindingLen, findings)
        End If

        ' Helper: token has visible (non-whitespace) chars
        Dim isVisibleToken As Func(Of Microsoft.Office.Interop.Word.Range, Boolean) =
    Function(w As Microsoft.Office.Interop.Word.Range)
        Dim t = w.Text
        Return Not String.IsNullOrEmpty(t) AndAlso t.Trim().Length > 0
    End Function

        ' 2) Aggregate additional heuristics into contiguous runs (emit only if span >= MinFindingLen)
        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        ProgressBarModule.GlobalProgressLabel = "Checking very small font size…"
        AddRuns(rng,
        Function(w)
            If Not isVisibleToken(w) Then Return False
            Return SafePt(w.Font.Size, 11.0F) < TinyFontPt
        End Function,
        "Very small font size",
        baseStart, MinFindingLen, findings)

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        ProgressBarModule.GlobalProgressLabel = "Checking font vs paragraph shading…"
        AddRuns(rng,
        Function(w)
            If Not isVisibleToken(w) Then Return False
            Return HasMeaningfulShading(w) AndAlso FontEqualsShadingColorIndex(w)
        End Function,
        "Font color equals background shading color",
        baseStart, MinFindingLen, findings)

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        ProgressBarModule.GlobalProgressLabel = "Checking white-on-white…"
        AddRuns(rng,
        Function(w)
            If Not isVisibleToken(w) Then Return False
            Return IsWhiteOnWhite(w)
        End Function,
        "Likely white-on-white (near-invisible) text",
        baseStart, MinFindingLen, findings)

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        ProgressBarModule.GlobalProgressLabel = "Checking font vs highlight color…"
        AddRuns(rng,
        Function(w)
            If Not isVisibleToken(w) Then Return False
            Return HasMeaningfulHighlight(w) AndAlso FontEqualsHighlightColorIndex(w)
        End Function,
        "Font color equals highlight color (camouflage)",
        baseStart, MinFindingLen, findings)

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        ProgressBarModule.GlobalProgressLabel = "Checking extreme font scaling…"
        AddRuns(rng,
        Function(w)
            If Not isVisibleToken(w) Then Return False
            Return SafePercent(w.Font.Scaling, 100) <= 10
        End Function,
        "Extreme font scaling (condensed)",
        baseStart, MinFindingLen, findings)

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        ProgressBarModule.GlobalProgressLabel = "Checking negative character spacing…"
        AddRuns(rng,
        Function(w)
            If Not isVisibleToken(w) Then Return False
            Return SafePercent(w.Font.Spacing, 0) < -2
        End Function,
        "Very negative character spacing",
        baseStart, MinFindingLen, findings)

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        ProgressBarModule.GlobalProgressLabel = "Checking font vs table cell shading…"
        AddRuns(rng,
        Function(w)
            Dim t = w.Text : If String.IsNullOrEmpty(t) OrElse t.Trim().Length = 0 Then Return False
            Return HasMeaningfulCellShading(w) AndAlso FontEqualsCellShadingColorIndex(w)
        End Function,
        "Font color equals table cell background color",
        baseStart, MinFindingLen, findings)

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
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

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
        ProgressBarModule.GlobalProgressLabel = "Checking zero-width/Bidi controls…"
        AddRuns(rng,
        Function(w)
            Dim raw As String = w.Text
            If String.IsNullOrEmpty(raw) Then Return False
            Return ContainsZeroWidthOrBidi(raw) AndAlso raw.Trim().Length > 0
        End Function,
        "Zero-width/Bidi control characters present",
        baseStart, MinFindingLen, findings)

        If ProgressBarModule.CancelOperation Then
            Return findings
        End If
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

        ' Complete & close progress (both normal/early exit handled)
        Try
            ProgressBarModule.GlobalProgressLabel = "Completed"
            ProgressBarModule.GlobalProgressValue = ProgressBarModule.GlobalProgressMax
            System.Threading.Tasks.Task.Delay(500)
            ProgressBarModule.CancelOperation = True
        Catch
        End Try

        Return result
    End Function


    ' Helper: slice a Range inside the SAME story as rng using absolute story coordinates
    Private Function SliceInSameStory(rng As Word.Range, absStart As Integer, absEnd As Integer) As Word.Range
        Dim slice As Word.Range = rng.Duplicate
        slice.Start = absStart
        slice.End = absEnd
        Return slice
    End Function

    ' Helper: slice by relative offsets from rng.Start inside the SAME story
    Private Function SliceByRel(rng As Word.Range, relStart As Integer, length As Integer) As Word.Range
        Dim absStart As Integer = rng.Start + System.Math.Max(0, relStart)
        Dim absEnd As Integer = absStart + System.Math.Max(0, length)
        Return SliceInSameStory(rng, absStart, absEnd)
    End Function

    ' Robust hidden-run detection by scanning characters and merging runs
    ' Also reveals hidden runs (unhide + red) when flushed.
    ' Progress: increments per character; cancellation honored.
    Private Sub AddHiddenRunsByFind(
    ByVal rng As Microsoft.Office.Interop.Word.Range,
    ByVal baseStart As Integer,
    ByVal minLen As Integer,
    ByVal findings As System.Collections.Generic.List(Of SuspiciousSpan)
)
        ' Fast skip: only proceed if range actually has any hidden formatting (True or wdUndefined)
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
                        inRun = True
                        runStartAbs = ch.Start
                        lastEndAbs = ch.End
                    Else
                        If ch.Start = lastEndAbs Then
                            lastEndAbs = ch.End
                        Else
                            Dim runLen As Integer = System.Math.Max(0, lastEndAbs - runStartAbs)
                            If runLen >= minLen Then
                                ' Reveal & record
                                RevealHiddenRun(rng, runStartAbs, lastEndAbs)
                                AddRunFinding(findings, "Hidden text span (revealed in red)", runStartAbs, lastEndAbs, baseStart, rng)
                            End If
                            runStartAbs = ch.Start
                            lastEndAbs = ch.End
                        End If
                    End If
                ElseIf inRun Then
                    Dim runLen As Integer = System.Math.Max(0, lastEndAbs - runStartAbs)
                    If runLen >= minLen Then
                        RevealHiddenRun(rng, runStartAbs, lastEndAbs)
                        AddRunFinding(findings, "Hidden text span (revealed in red)", runStartAbs, lastEndAbs, baseStart, rng)
                    End If
                    inRun = False
                    runStartAbs = -1
                    lastEndAbs = -1
                End If

                ' progress tick per character
                stepCounter += 1
                Try
                    ProgressBarModule.GlobalProgressValue = System.Math.Min(ProgressBarModule.GlobalProgressValue + 1, ProgressBarModule.GlobalProgressMax)
                    If (stepCounter Mod 500) = 0 Then
                        ProgressBarModule.GlobalProgressLabel = $"Scanning hidden text… ({stepCounter}/{System.Math.Max(1, chars.Count)})"
                    End If
                Catch
                End Try
            Next

            ' flush trailing run
            If inRun AndAlso runStartAbs >= 0 AndAlso lastEndAbs > runStartAbs Then
                Dim runLen As Integer = System.Math.Max(0, lastEndAbs - runStartAbs)
                If runLen >= minLen Then
                    RevealHiddenRun(rng, runStartAbs, lastEndAbs)
                    AddRunFinding(findings, "Hidden text span (revealed in red)", runStartAbs, lastEndAbs, baseStart, rng)
                End If
            End If
        Catch
            ' best effort
        End Try
    End Sub

    ' Unhide and color a hidden span in red without altering text content/positions
    Private Sub RevealHiddenRun(ByVal selectionRange As Microsoft.Office.Interop.Word.Range,
                            ByVal startAbs As Integer,
                            ByVal endAbs As Integer)
        Try
            Dim dr As Microsoft.Office.Interop.Word.Range = SliceInSameStory(selectionRange, startAbs, endAbs)
            dr.Font.Hidden = 0
            dr.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdRed
        Catch
            ' ignore formatting failures
        End Try
    End Sub

    ' Aggregates contiguous Word ranges satisfying predicate into a single finding per run
    ' Progress: increments per word; cancellation honored.
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

            ' progress tick per word
            Try
                ProgressBarModule.GlobalProgressValue = System.Math.Min(ProgressBarModule.GlobalProgressValue + 1, ProgressBarModule.GlobalProgressMax)
            Catch
            End Try
        Next

        If inRun Then
            Dim runLen As Integer = System.Math.Max(0, runEndAbs - runStartAbs)
            If runLen >= minLen Then
                AddRunFinding(findings, reason, runStartAbs, runEndAbs, baseStart, rng)
            End If
        End If
    End Sub

    ' ===== Extend camouflage detection =====

    ' Detect white-on-white also when Font.ColorIndex is Auto but resolved RGB is near-white
    ' (keeps your existing logic; additional AddRuns not needed as IsWhiteOnWhite already used)
    Private Function IsWhiteOnWhite(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            Dim fci = w.Font.ColorIndex
            Dim fontIsWhiteIdx As Boolean = (fci = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite)

            ' If Auto/ByAuthor, fall back to explicit RGB check
            Dim rgb As Integer = SafeWdColorToRgb(w.Font.Color)
            Dim nearWhite As Boolean = IsLikelyInvisibleOnWhite(rgb)

            Dim hasShade As Boolean = HasMeaningfulShading(w)
            Dim hasHl As Boolean = HasMeaningfulHighlight(w)

            If fontIsWhiteIdx Then
                If Not hasShade AndAlso Not hasHl Then Return True
                If hasHl AndAlso w.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite Then Return True
                If hasShade AndAlso w.Shading.BackgroundPatternColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdWhite Then Return True
            End If

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

    ' Consider table-cell shading too (camouflage inside tables)
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


    ' Build finding using SAME-story slice to avoid main-story misalignment
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

    ' Build finding using Word's Range slice to avoid index misalignment (hidden/control chars)
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
            ' Use SAME-story slice for correctness across stories
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

    ' Only consider shading if it's explicitly set (not Auto/None/ByAuthor), then compare ColorIndex
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

    Private Function HasMeaningfulHighlight(w As Microsoft.Office.Interop.Word.Range) As Boolean
        Try
            Dim hi = w.HighlightColorIndex
            Return hi <> Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight AndAlso
                   hi <> Microsoft.Office.Interop.Word.WdColorIndex.wdByAuthor
        Catch
            Return False
        End Try
    End Function

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



    ' Aggregates contiguous Word ranges satisfying predicate into a single finding per run
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


    Private Function IsLikelyInvisibleOnWhite(ByVal rgb As System.Int32) As System.Boolean
        If rgb = -1 Then Return False
        Dim r As System.Int32 = (rgb And &HFF)
        Dim g As System.Int32 = ((rgb >> 8) And &HFF)
        Dim b As System.Int32 = ((rgb >> 16) And &HFF)
        Dim avg As System.Int32 = (r + g + b) \ 3
        Return avg >= 245
    End Function

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



    Private Function SafePt(ByVal size As System.Object, ByVal fallbackPt As System.Single) As System.Single
        Try
            Return System.Convert.ToSingle(size, Globalization.CultureInfo.InvariantCulture)
        Catch
            Return fallbackPt
        End Try
    End Function

    Private Function SafePercent(ByVal val As System.Object, ByVal fallback As System.Int32) As System.Int32
        Try
            Return System.Convert.ToInt32(val, Globalization.CultureInfo.InvariantCulture)
        Catch
            Return fallback
        End Try
    End Function

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


    Private NotInheritable Class SuspiciousSpan
        Public Property Reason As System.String
        Public Property StartIndex As System.Int32
        Public Property Length As System.Int32
        Public Property Snippet As System.String
    End Class

End Class
