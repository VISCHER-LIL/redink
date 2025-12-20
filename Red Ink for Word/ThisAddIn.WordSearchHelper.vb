' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Explicit On
Option Strict Off

Public Module WordSearchHelper

    Private deletionsCache As System.Collections.Generic.List(Of (Integer, Integer)) = Nothing

    Private ReadOnly RX_U As New _
    System.Text.RegularExpressions.Regex("\\u([0-9A-Fa-f]{4,6})",
        System.Text.RegularExpressions.RegexOptions.Compiled Or
        System.Text.RegularExpressions.RegexOptions.CultureInvariant)

    Private ReadOnly RX_ELL As New _
    System.Text.RegularExpressions.Regex("\.\.\.",
        System.Text.RegularExpressions.RegexOptions.Compiled Or
        System.Text.RegularExpressions.RegexOptions.CultureInvariant)

    ' Toggle detailed per-slice logging if needed
    Private Const ENABLE_SLICE_DEBUG As Boolean = True

    Public Function FindLongTextAnchoredFast(
        ByRef sel As Microsoft.Office.Interop.Word.Selection,
        ByVal findText As System.String,
        Optional ByVal skipDeleted As System.Boolean = True,
        Optional ByVal nWords As System.Int32 = 4,
        Optional ByVal cancel As System.Threading.CancellationToken = Nothing,
        Optional ByVal timeoutSeconds As System.Int32 = 10
    ) As System.Boolean

        'System.Diagnostics.Debug.WriteLine("Skipdeleted=" & skipDeleted)

        Dim wordApp As Microsoft.Office.Interop.Word.Application = sel.Application

        ' Preserve and only change/restore view if required
        Dim view = wordApp.ActiveWindow.View
        Dim origRevView = view.RevisionsView
        Dim origShowRev = view.ShowRevisionsAndComments
        Dim viewChanged1 As Boolean = False
        Dim viewChanged2 As Boolean = False

        If skipDeleted Then
            If view.RevisionsView <> Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal Then
                view.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                viewChanged1 = True
            End If
            If view.ShowRevisionsAndComments Then
                view.ShowRevisionsAndComments = False
                viewChanged2 = True
            End If
        End If

        Dim timedOut As Boolean = False

        Try
            Dim _dbgLastSlice As System.String = ""
            Dim _dbgLastIdx As System.Int32 = -1

            Dim t0 As System.DateTime = System.DateTime.UtcNow

            Dim doc As Microsoft.Office.Interop.Word.Document = sel.Document
            Dim mainStory As Microsoft.Office.Interop.Word.Range =
                doc.StoryRanges(Microsoft.Office.Interop.Word.WdStoryType.wdMainTextStory).Duplicate

            Dim area As Microsoft.Office.Interop.Word.Range
            If sel.Range.Start = sel.Range.End Then
                area = mainStory.Duplicate
            Else
                Dim sStart As System.Int32 = System.Math.Max(sel.Range.Start, mainStory.Start)
                Dim sEnd As System.Int32 = System.Math.Min(sel.Range.End, mainStory.End)
                If sEnd < sStart Then sEnd = sStart
                area = doc.Range(Start:=sStart, End:=sEnd)
            End If

            ' 0) Plain literal
            If findText.Length <= 255 Then
                Dim rngPlain As Microsoft.Office.Interop.Word.Range = area.Duplicate
                With rngPlain.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = findText
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWholeWord = False
                    .MatchWildcards = False : .Format = False
                    .IgnoreSpace = False
                End With
                Dim hitPlain As System.Boolean
                Try : hitPlain = rngPlain.Find.Execute() : Catch : hitPlain = False : End Try
                If hitPlain Then
                    sel.SetRange(rngPlain.Start, rngPlain.End)
                    Return True
                End If
            End If

            ' 1) Masked literal wildcard (IgnoreSpace)
            Dim litPat As System.String = EscapeForWordWildcard(findText)
            If litPat.Length <= 255 Then
                Dim rngLit As Microsoft.Office.Interop.Word.Range = area.Duplicate
                With rngLit.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = litPat
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWildcards = True
                    .Format = False : .IgnoreSpace = True
                End With
                If rngLit.Find.Execute() Then
                    sel.SetRange(rngLit.Start, rngLit.End)
                    Return True
                End If
            End If

            ' 2) Needle prep
            Dim raw() As System.String = WordSearchHelper.RawWords(findText)
            Dim canonNeedle As System.String = Canonicalise(findText, True)
            If canonNeedle.Length = 0 Then
                ' Nothing to search for
                Return False
            End If

            ' Try full wildcard if short enough
            Dim fullWildcardPattern As System.String = BuildWildcardProbe(raw)
            If fullWildcardPattern.Length <= 255 Then
                Dim rngFull As Microsoft.Office.Interop.Word.Range = area.Duplicate
                With rngFull.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = fullWildcardPattern
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWildcards = True
                    .Format = False : .IgnoreSpace = True
                End With
                If rngFull.Find.Execute() Then
                    sel.SetRange(rngFull.Start, rngFull.End)
                    Return True
                End If
            End If

            If raw.Length < 2 Then Return False

            nWords = System.Math.Min(nWords, raw.Length \ 2)
            If nWords < 1 Then nWords = 1

            Do While nWords > 1 AndAlso BuildWildcardProbe(raw.Take(nWords).ToArray()).Length > 255
                nWords -= 1
            Loop

            Dim startPat As System.String = BuildWildcardProbe(raw.Take(nWords).ToArray())
            Dim endWords() As System.String = raw.Skip(raw.Length - nWords).ToArray()
            Dim endPat As System.String = BuildWildcardProbe(endWords)

            Dim occur As System.Int32 = CountOccurrences(findText, System.String.Join(" "c, endWords))
            If startPat = endPat Then occur = System.Math.Max(2, occur)

            deletionsCache = Nothing

            ' 3) Anchored search
            Using sRng As New RangeProxy(area.Duplicate)
                Dim fS As Microsoft.Office.Interop.Word.Find = sRng.Range.Find
                With fS
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Font.Reset() : .ParagraphFormat.Reset()
                    .Text = startPat
                    .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                    .MatchCase = False : .MatchWildcards = True
                    .Format = False : .IgnoreSpace = True
                End With

                Dim okS As System.Boolean : Try : okS = fS.Execute() : Catch : okS = False : End Try
                While okS
                    If (System.DateTime.UtcNow - t0).TotalSeconds > timeoutSeconds Then
                        timedOut = True : Exit While
                    End If
                    cancel.ThrowIfCancellationRequested()

                    Dim posStart As System.Int32 = sRng.Range.Start
                    Dim searchFrom As System.Int32 = sRng.Range.End

                    Dim eRng As Microsoft.Office.Interop.Word.Range = doc.Range(searchFrom, area.End)
                    Dim fE As Microsoft.Office.Interop.Word.Find = eRng.Find
                    With fE
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Font.Reset() : .ParagraphFormat.Reset()
                        .Text = endPat
                        .Forward = True : .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                        .MatchCase = False : .MatchWildcards = True
                        .Format = False : .IgnoreSpace = True
                    End With
                    Dim okE As System.Boolean : Try : okE = fE.Execute() : Catch : okE = False : End Try
                    For i As System.Int32 = 2 To occur
                        If Not okE Then Exit For
                        eRng.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        Try : okE = fE.Execute() : Catch : okE = False : End Try
                    Next

                    If okE Then
                        Dim sliceTxt As System.String
                        Dim back As System.Collections.Generic.IReadOnlyList(Of System.Int32)
                        VisibleSlice(doc, posStart, eRng.End - posStart, skipDeleted, sliceTxt, back)

                        If ENABLE_SLICE_DEBUG AndAlso System.Diagnostics.Debugger.IsAttached Then
                            System.Diagnostics.Debug.WriteLine(sliceTxt & System.Environment.NewLine)
                        End If

                        Dim canSlice As System.String
                        Dim backCanon As System.Collections.Generic.List(Of System.Int32)
                        CanonicaliseWithBackMap(sliceTxt, True, back, canSlice, backCanon)

                        Dim idx As System.Int32 = canSlice.IndexOf(canonNeedle, System.StringComparison.Ordinal)
                        _dbgLastSlice = canSlice
                        _dbgLastIdx = idx


                        If idx >= 0 Then
                            Dim endIdx As System.Int32 = System.Math.Min(idx + canonNeedle.Length - 1, backCanon.Count - 1)

                            If System.Diagnostics.Debugger.IsAttached Then
                                For z As Integer = 1 To backCanon.Count - 1
                                    If backCanon(z) < backCanon(z - 1) Then
                                        System.Diagnostics.Debug.WriteLine(
                                                    "BACKMAP BREAK: index=" & z &
                                                    "  prev=" & backCanon(z - 1) &
                                                    "  curr=" & backCanon(z) &
                                                    "  Δ=" & (backCanon(z) - backCanon(z - 1))
                                                )
                                        Exit For
                                    End If
                                Next

                                System.Diagnostics.Debug.WriteLine("HIT idx=" & idx &
                                        " maps to Start=" & backCanon(idx) &
                                        " End=" & backCanon(endIdx))
                            End If

                            sel.SetRange(backCanon(idx), backCanon(endIdx) + 1)
                            Return True
                        End If


                    End If

                    sRng.CollapseEndPlusOne()
                    If sRng.Range.Start >= area.End Then Exit While
                    Try : okS = fS.Execute() : Catch : okS = False : End Try
                End While
            End Using

            ' 4) Windowed fallback (canon)
            Dim winSize As System.Int32 = 12000
            Dim overlap As System.Int32 = 400
            Dim p As System.Int32 = area.Start
            While p < area.End
                If (System.DateTime.UtcNow - t0).TotalSeconds > timeoutSeconds Then
                    timedOut = True : Exit While
                End If
                cancel.ThrowIfCancellationRequested()

                Dim len As System.Int32 = System.Math.Min(winSize, area.End - p)
                Dim sliceTxt As System.String
                Dim back As System.Collections.Generic.IReadOnlyList(Of System.Int32)
                VisibleSlice(doc, p, len, skipDeleted, sliceTxt, back)

                Dim canSlice As System.String
                Dim backCanon As System.Collections.Generic.List(Of System.Int32)
                CanonicaliseWithBackMap(sliceTxt, True, back, canSlice, backCanon)

                Dim idx As System.Int32 = canSlice.IndexOf(canonNeedle, System.StringComparison.Ordinal)

                ' Temporary replace block for production use xxxxxx

                If idx >= 0 Then
                    Dim endIdx As System.Int32 = System.Math.Min(idx + canonNeedle.Length - 1, backCanon.Count - 1)

                    ' Get the mapped positions with boundary checking
                    Dim selStart As System.Int32 = backCanon(idx)
                    Dim selEnd As System.Int32 = backCanon(endIdx)

                    ' Ensure we don't exceed document boundaries
                    selStart = System.Math.Max(selStart, doc.Content.Start)
                    selEnd = System.Math.Min(selEnd + 1, doc.Content.End)

                    ' Verify valid range
                    If selEnd <= selStart Then
                        Continue While  ' Skip this match and continue searching
                    End If

                    Try
                        Dim testRange As Microsoft.Office.Interop.Word.Range = doc.Range(selStart, selEnd)

                        ' Verify the selection contains the expected text length
                        While testRange.Text.Length < findText.Length AndAlso testRange.End < doc.Content.End - 1
                            testRange.End = testRange.End + 1
                        End While

                        ' Final boundary check before setting selection
                        If testRange.Start >= doc.Content.Start AndAlso testRange.End <= doc.Content.End Then
                            sel.SetRange(testRange.Start, testRange.End)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(testRange)
                            Return True
                        End If

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(testRange)
                    Catch ex As System.Runtime.InteropServices.COMException
                        ' Handle range errors - continue searching
                        If System.Diagnostics.Debugger.IsAttached Then
                            System.Diagnostics.Debug.WriteLine($"Range error at Start={selStart}, End={selEnd}: {ex.Message}")
                        End If
                    End Try
                End If

                GoTo SkipToNext

                ' Temporary replace block for debugging xxxxxx

                If idx >= 0 Then
                    Dim endIdx As System.Int32 = System.Math.Min(idx + canonNeedle.Length - 1, backCanon.Count - 1)

                    ' Get the mapped positions
                    Dim selStart As System.Int32 = backCanon(idx)
                    Dim selEnd As System.Int32 = backCanon(endIdx)

                    ' Adjust end position to include the last character
                    ' Word ranges are inclusive of start but exclusive of end
                    Dim testRange As Microsoft.Office.Interop.Word.Range = doc.Range(selStart, selEnd + 1)

                    ' Verify the selection contains the expected text length
                    While testRange.Text.Length < findText.Length AndAlso testRange.End < doc.Content.End
                        testRange.End = testRange.End + 1
                    End While

                    sel.SetRange(testRange.Start, testRange.End)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(testRange)
                    Return True
                End If

SkipToNext:

                p += System.Math.Max(1, winSize - overlap)
            End While

            ' FINAL DEBUG now printed only when we actually give up
            If System.Diagnostics.Debugger.IsAttached Then
                Dim sliceLen As Integer = If(_dbgLastSlice Is Nothing, 0, _dbgLastSlice.Length)
                Dim containsStr As String
                If sliceLen = 0 OrElse canonNeedle.Length = 0 Then
                    containsStr = "n/a"
                Else
                    containsStr = _dbgLastSlice.Contains(canonNeedle).ToString()
                End If

                System.Diagnostics.Debug.WriteLine("===== FindLongTextAnchoredFast: FINAL DEBUG =====")
                System.Diagnostics.Debug.WriteLine("  findText        = '" & findText & "'")
                System.Diagnostics.Debug.WriteLine("  lastIdx         = " & _dbgLastIdx)
                System.Diagnostics.Debug.WriteLine("  needle.Length   = " & canonNeedle.Length)
                System.Diagnostics.Debug.WriteLine("  slice.Length    = " & sliceLen)
                System.Diagnostics.Debug.WriteLine("  contains?       = " & containsStr)
                Dim previewLen As System.Int32 = 200
                Dim startEx As System.String = If(sliceLen <= previewLen,
                                              _dbgLastSlice,
                                              _dbgLastSlice.Substring(0, previewLen) & "…")
                Dim endEx As System.String = If(sliceLen <= previewLen, "",
                                            "…" & _dbgLastSlice.Substring(sliceLen - previewLen))
                System.Diagnostics.Debug.WriteLine("  slice excerpt start: '" & startEx & "'")
                If endEx <> "" Then
                    System.Diagnostics.Debug.WriteLine("  slice excerpt end:   '" & endEx & "'")
                End If
                If timedOut Then
                    System.Diagnostics.Debug.WriteLine("  NOTE: search aborted due to timeout.")
                End If
                System.Diagnostics.Debug.WriteLine("===============================================")
            End If

            Return False

        Finally
            If viewChanged1 Then view.RevisionsView = origRevView
            If viewChanged2 Then view.ShowRevisionsAndComments = origShowRev
            ' IMPORTANT: do not ReleaseComObject(wordApp) here (root object)
        End Try
    End Function




    Private Sub CanonicaliseWithBackMap(
        ByVal src As String,
        ByVal collapseWS As Boolean,
        ByVal backIn As System.Collections.Generic.IReadOnlyList(Of Integer),
        ByRef canonOut As String,
        ByRef canonBack As System.Collections.Generic.List(Of Integer))

        src = PrepareNeedle(src).Normalize(System.Text.NormalizationForm.FormKC)

        Dim sb As New System.Text.StringBuilder(src.Length)
        Dim back As New System.Collections.Generic.List(Of Integer)(src.Length)
        Dim pendingSpace As Boolean = False

        For i As Integer = 0 To src.Length - 1
            Dim ch As Char = src(i)
            If IsDocNoise(ch) Then Continue For

            Dim code As Integer = AscW(ch)
            Dim isHyphenOrWs As Boolean = System.Char.IsWhiteSpace(ch) OrElse code = &HA0
            If Not isHyphenOrWs Then
                Select Case code
                    Case &H2010, &H2011, &H2013, &H2014, &HAD, 45
                        isHyphenOrWs = True
                End Select
            End If

            If isHyphenOrWs Then
                pendingSpace = True
            Else
                If pendingSpace AndAlso collapseWS Then
                    sb.Append(" "c)
                    'Dim mapIdx As Integer = System.Math.Min(System.Math.Max(i - 1, 0), backIn.Count - 1)
                    Dim mapIdx As Integer = System.Math.Min(i, System.Math.Max(backIn.Count - 1, 0))
                    back.Add(If(backIn.Count > 0, backIn(mapIdx), 0))
                End If
                pendingSpace = False

                Select Case AscW(ch)
                    Case &HDF, &H1E9E
                        sb.Append("S"c) : sb.Append("S"c)
                        Dim mi As Integer = System.Math.Min(i, System.Math.Max(backIn.Count - 1, 0))
                        Dim m As Integer = If(backIn.Count > 0, backIn(mi), 0)
                        back.Add(m) : back.Add(m)
                    Case Else
                        Dim up As Char = System.Char.ToUpperInvariant(ch)
                        sb.Append(up)
                        Dim mi As Integer = System.Math.Min(i, System.Math.Max(backIn.Count - 1, 0))
                        back.Add(If(backIn.Count > 0, backIn(mi), 0))
                End Select
            End If
        Next

        Dim s As String = sb.ToString()
        Dim start As Integer = 0
        While start < s.Length AndAlso System.Char.IsWhiteSpace(s.Chars(start))
            start += 1
        End While
        Dim [end] As Integer = s.Length
        While [end] > start AndAlso System.Char.IsWhiteSpace(s.Chars([end] - 1))
            [end] -= 1
        End While

        canonOut = If([end] > start, s.Substring(start, [end] - start), System.String.Empty)
        canonBack = If([end] > start, back.GetRange(start, [end] - start), New System.Collections.Generic.List(Of Integer)())

        If System.Diagnostics.Debugger.IsAttached Then
            System.Diagnostics.Debug.WriteLine("canonOut=""" & canonOut & """")
            System.Diagnostics.Debug.WriteLine("canonBack length=" & canonBack.Count)
        End If

    End Sub

    Private Function BuildWildcardProbe(ByVal words() As String) As String
        Dim sb As New System.Text.StringBuilder(words.Length * 14)
        Dim i As Integer = 0
        While i < words.Length
            If i > 0 Then sb.Append(" "c)

            Dim w As String = words(i)
            If w.Contains("["c) Then
                While i < words.Length AndAlso Not words(i).Contains("]"c)
                    i += 1
                End While
                sb.Append("\[*\]")
                If i < words.Length Then
                    Dim rest As String = words(i).Substring(words(i).IndexOf("]"c) + 1)
                    If rest <> "" Then sb.Append(EscapeForWordWildcard(rest))
                End If
            Else
                w = w.Replace("-"c, "?"c) _
                     .Replace(ChrW(&H2010), "?"c) _
                     .Replace(ChrW(&H2011), "?"c) _
                     .Replace(ChrW(&H2013), "?"c) _
                     .Replace(ChrW(&H2014), "?"c) _
                     .Replace(ChrW(&HAD), "?"c)
                sb.Append(EscapeForWordWildcard(w))
            End If
            i += 1
        End While
        Return sb.ToString()
    End Function

    Private Function CountOccurrences(ByVal txt As String, ByVal subTxt As String) As Integer
        txt = Canonicalise(txt, True)
        subTxt = Canonicalise(subTxt, True)
        Dim cnt As Integer = 0
        Dim pos As Integer = txt.IndexOf(subTxt, System.StringComparison.OrdinalIgnoreCase)
        While pos <> -1
            cnt += 1
            pos = txt.IndexOf(subTxt, pos + subTxt.Length, System.StringComparison.OrdinalIgnoreCase)
        End While
        Return cnt
    End Function

    Private Function RawWords(ByVal src As String) As String()
        src = RX_U.Replace(src, Function(m) _
            System.Char.ConvertFromUtf32(System.Convert.ToInt32(m.Groups(1).Value, 16)))
        src = RX_ELL.Replace(src, ChrW(&H2026))
        src = src.Normalize(System.Text.NormalizationForm.FormKC)
        Return src.Split(New Char() {" "c, ChrW(9), ChrW(10), ChrW(13)},
                         System.StringSplitOptions.RemoveEmptyEntries)
    End Function

    Private Sub VisibleSlice(
    ByVal doc As Microsoft.Office.Interop.Word.Document,
    ByVal absStart As System.Int32,
    ByVal sliceLen As System.Int32,
    ByVal skipDeleted As System.Boolean,
    ByRef visOut As System.String,
    ByRef mapBack As System.Collections.Generic.IReadOnlyList(Of System.Int32))

        ' Build raw window once with a small safety margin
        Dim rawEnd As System.Int32 = System.Math.Min(doc.Content.End, absStart + sliceLen + 500)
        Dim rawRng As Microsoft.Office.Interop.Word.Range = doc.Range(absStart, rawEnd)
        Dim rawTxt As System.String = rawRng.Text
        Dim rawLen As System.Int32 = rawTxt.Length

        Dim take As System.Int32 = System.Math.Min(sliceLen, rawLen)
        If take < 0 Then take = 0
        visOut = If(rawLen > take, rawTxt.Substring(0, take), rawTxt)

        ' Build proper position mapping
        Dim backList As New System.Collections.Generic.List(Of System.Int32)(visOut.Length)

        If visOut.Length > 0 Then
            ' Create a temporary range for accurate position mapping
            Dim tempRng As Microsoft.Office.Interop.Word.Range = doc.Range(absStart, absStart)
            Dim currentPos As System.Int32 = absStart

            For i As System.Int32 = 0 To visOut.Length - 1
                ' Map string position to actual Word position
                backList.Add(currentPos)

                ' Move to next character position in the document
                If i < visOut.Length - 1 Then
                    Try
                        tempRng.SetRange(currentPos, rawEnd)
                        tempRng.MoveStart(Microsoft.Office.Interop.Word.WdUnits.wdCharacter, 1)
                        currentPos = tempRng.Start
                    Catch
                        ' Fallback to linear increment if range operations fail
                        currentPos += 1
                    End Try
                End If
            Next

            System.Runtime.InteropServices.Marshal.ReleaseComObject(tempRng)
        End If

        mapBack = backList
    End Sub


    Private Function PrepareNeedle(ByVal txt As String) As String
        txt = RX_U.Replace(txt,
            Function(m) System.Char.ConvertFromUtf32(
                System.Convert.ToInt32(m.Groups(1).Value, 16)))
        Return RX_ELL.Replace(txt, ChrW(&H2026))
    End Function

    Private Function Canonicalise(ByVal src As String, ByVal collapseWS As Boolean) As String
        src = PrepareNeedle(src).Normalize(System.Text.NormalizationForm.FormKC)

        Dim sb As New System.Text.StringBuilder(src.Length)
        Dim pendingSpace As Boolean = False

        For Each ch As Char In src
            If IsDocNoise(ch) Then Continue For

            Dim code As Integer = AscW(ch)
            Dim isHyphenOrWs As Boolean = System.Char.IsWhiteSpace(ch) OrElse code = &HA0
            If Not isHyphenOrWs Then
                Select Case code
                    Case &H2010, &H2011, &H2013, &H2014, &HAD, 45
                        isHyphenOrWs = True
                End Select
            End If

            If isHyphenOrWs Then
                pendingSpace = True
            Else
                If pendingSpace AndAlso collapseWS Then sb.Append(" "c)
                pendingSpace = False
                sb.Append(CanonizeDocChar(ch))
            End If
        Next
        Return sb.ToString().Trim()
    End Function

    Private Function IsDocNoise(ByVal ch As Char) As Boolean
        Dim code As Integer = AscW(ch)
        If code < 32 AndAlso code <> 9 AndAlso code <> 10 AndAlso code <> 13 Then Return True
        Select Case code
            Case &HA0, &H200B, &H200C, &H200D, &H2060,
                 &H200E To &H200F, &H202A To &H202E,
                 1, 19, 20, 21, &HFFFA, &HFFFB, &HFFFC
                Return True
        End Select
        Return False
    End Function

    Private Function CanonizeDocChar(ByVal ch As Char) As String
        Select Case AscW(ch)
            Case &HDF, &H1E9E : Return "SS"
            Case Else : Return System.Char.ToUpperInvariant(ch)
        End Select
    End Function

    Private Function EscapeForWordWildcard(ByVal s As String) As String
        If s = "" Then Return ""
        Dim sb As New System.Text.StringBuilder(s.Length * 2)
        For Each ch As Char In s
            Select Case ch
                Case "?"c, "*"c, "@"c, "["c, "]"c, "("c, ")"c,
                     "{"c, "}"c, "\"c, "<"c, ">"c
                    sb.Append("\"c)
            End Select
            sb.Append(ch)
        Next
        Return sb.ToString()
    End Function

    Private NotInheritable Class RangeProxy
        Implements System.IDisposable

        Friend ReadOnly Range As Microsoft.Office.Interop.Word.Range
        Private ReadOnly ptr As Object

        Friend Sub New(ByVal r As Microsoft.Office.Interop.Word.Range)
            Range = r
            ptr = r
        End Sub

        Friend Sub CollapseEndPlusOne()
            Range.Collapse(
                Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            Range.SetRange(Range.Start + 1, Range.Start + 1)
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If ptr IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ptr)
            End If
        End Sub
    End Class

End Module

