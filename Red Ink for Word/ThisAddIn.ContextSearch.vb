' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict Off

Imports System.Data
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NAudio.Wave
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn


    Private embed_store As EmbeddingStore
    Private embed_indexedDocs As HashSet(Of String) = New HashSet(Of String)()

    Public Async Sub ContextSearch()

        Dim Prefix As String = "-CS"

        If INILoadFail() Then Return

        Dim EmbedModel As String = ""
        Dim EmbedVocab As String = ""

        If Not String.IsNullOrEmpty(INI_LocalModelPath) Then

            EmbedModel = System.IO.Path.Combine(ExpandEnvironmentVariables(INI_LocalModelPath), Embed_Model)
            EmbedVocab = System.IO.Path.Combine(ExpandEnvironmentVariables(INI_LocalModelPath), Embed_Vocab)

            If File.Exists(EmbedModel) And File.Exists(EmbedVocab) Then
                If embed_store Is Nothing Then embed_store = New EmbeddingStore(EmbedModel, EmbedVocab)
            End If
        End If

        Dim EmbeddingAvailable As Boolean = Not embed_store Is Nothing

        Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.ActiveDocument
        Dim DoSearchNext As Boolean = False
        Dim EmbedInstruct As String = If(EmbeddingAvailable, $"add '{EmbedTrigger} to use embeddings, ", "")
        Dim DoBoW As Boolean = False
        Dim DoRefresh As Boolean = False
        Dim DoEmbed As Boolean = False

        Dim lastcontextsearch As String = If(String.IsNullOrWhiteSpace(My.Settings.LastContextSearch), "", My.Settings.LastContextSearch)

        SearchContext = ShowCustomInputBox($"Enter the search term (use '{SearchNextTrigger}' if you only want to find the next term; {EmbedInstruct}'{BoWTrigger}' to use Bag of Words and '{RefreshTrigger}' to refresh the index first):", "Context Search", True, lastcontextsearch).Trim()
        If String.IsNullOrWhiteSpace(SearchContext) Or SearchContext = "ESC" Then Return

        My.Settings.LastContextSearch = SearchContext
        My.Settings.Save()

        If SearchContext.StartsWith(SearchNextTrigger, StringComparison.OrdinalIgnoreCase) Then
            SearchContext = SearchContext.Substring(SearchNextTrigger.Length).Trim()
            DoSearchNext = True
        End If

        If SearchContext.IndexOf(EmbedTrigger, StringComparison.OrdinalIgnoreCase) >= 0 And EmbeddingAvailable Then
            SearchContext = SearchContext.Replace(EmbedTrigger, "").Trim()
            DoEmbed = True
        ElseIf SearchContext.IndexOf(BoWTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            SearchContext = SearchContext.Replace(BoWTrigger, "").Trim()
            DoBoW = True
        End If
        If SearchContext.IndexOf(RefreshTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            SearchContext = SearchContext.Replace(RefreshTrigger, "").Trim()
            DoRefresh = True
        End If

        SearchContext = SearchContext.Replace("  ", "")

        If DoEmbed Then
            RunSearch_Embed(EmbedModel, EmbedVocab, SearchContext, DoSearchNext, DoRefresh)
            Return
        ElseIf DoBoW Then
            RunSearch_bow(SearchContext, DoSearchNext, DoRefresh)
            Return
        End If

        Dim SearchText As String = ""

        If Not String.IsNullOrWhiteSpace(selection.Text) And Len(selection.Text) > 3 And DoSearchNext Then
            SearchText = selection.Text
        ElseIf selection.Start < selection.Document.Content.End And DoSearchNext Then
            SearchText = selection.Document.Range(selection.Start, selection.Document.Content.End).Text
            selection.SetRange(selection.Start, selection.Document.Content.End)
        Else
            SearchText = selection.Document.Content.Text
            selection.SetRange(0, selection.Document.Content.End)
            DoSearchNext = False
        End If

        Dim LLMResult As String = Await LLM(InterpolateAtRuntime(If(DoSearchNext, SP_ContextSearch, SP_ContextSearchMulti)), "<TEXTTOSEARCH>" & SearchText & "</TEXTTOSEARCH>", "", "", 0)

        LLMResult = LLMResult.Replace("<TEXTTOSEARCH>", "").Replace("</TEXTTOSEARCH>", "")

        If Not DoSearchNext Then

            Dim parts() As String = LLMResult.Split(New String() {"@@@"}, StringSplitOptions.RemoveEmptyEntries)
            Dim notFoundParts As New List(Of String)
            Dim originalStart As Integer = selection.Start
            Dim originalEnd As Integer = selection.End

            If parts.Count > 0 Then

                Dim splash As New Slib.Splashscreen($"Highlighting {parts.Count} hits... Press 'Esc' to abort")
                splash.Show()
                splash.Refresh()

                Dim Aborted As Boolean = False

                Dim trackChangesEnabled As Boolean = doc.TrackRevisions
                Dim originalAuthor As String = doc.Application.UserName

                doc.TrackRevisions = True

                Dim SuccessHits As Integer = 0

                For Each part As String In parts

                    splash.UpdateMessage($"Highlighting {parts.Count - SuccessHits} hits... Press 'Esc' to abort")

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
                        Aborted = True
                        Exit For
                    End If

                    If (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Then
                        ' Exit the loop
                        Aborted = True
                        Exit For
                    End If

                    Dim findText As String = part.Trim()
                    If FindLongTextInChunks(findText, selection) And selection IsNot Nothing Then
                        'selection.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow
                        doc.Comments.Add(selection.Range, $"{AN5}{Prefix}: '{SearchContext}'")
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        SuccessHits += 1
                    Else
                        notFoundParts.Add(findText)
                    End If
                Next

                splash.Close()

                If Aborted Then
                    ShowCustomMessageBox($"Search aborted. {SuccessHits} hit(s) have been highlighted so far.", "Context Search")
                ElseIf notFoundParts.Count > 0 Then
                    Dim errorlist As String = ShowCustomWindow($"{SuccessHits} hit(s) have been highlighted using Context Search. The following hit(s) could not be found:", String.Join(vbCrLf, notFoundParts), "The above error list will be included in a final comment at the end of your last hit (it will also be included in the clipboard). You can have the original list included, or you can now make changes and have this version used. If you select Cancel, nothing will be put added to the document.", AN, True)
                    If errorlist <> "" And errorlist.ToLower() <> "esc" Then
                        SLib.PutInClipboard(errorlist)
                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(selection.Range, $"{AN5} could not locate these sections: " & vbCrLf & errorlist)
                    End If
                Else
                    ShowCustomMessageBox($"{SuccessHits} hit(s) have been highlighted.", "Context Search")
                End If

                ' Restore the original selection
                selection.SetRange(originalStart, originalEnd)
                doc.TrackRevisions = trackChangesEnabled

            Else
                ShowCustomMessageBox($"The LLM has found no hits for the context '{SearchContext}'.", "Context Search")
            End If

        Else
            If Not String.IsNullOrWhiteSpace(LLMResult) Then
                Dim FindText As String = LLMResult.Trim()

                If FindLongTextInChunks(FindText, selection) And selection IsNot Nothing Then
                    wordApp.ActiveWindow.ScrollIntoView(selection.Range, True)
                Else
                    ShowCustomMessageBox($"The LLM found this section:" & vbCrLf & vbCrLf & FindText & vbCrLf & vbCrLf & $"However, {AN} could not locate it in the document for technical reasons (may be due to special characters, line breaks of the LLM not quoting the text properly).", "Context Search")
                End If
            Else
                ShowCustomMessageBox($"The LLM did not find any (further) hits for the context '{SearchContext}'.", "Context Search")
            End If
        End If
    End Sub


    Public Sub RunIndexing_Embed(refresh As Boolean, EmbedModel As String, EmbedVocab As String, ChunkLength As Integer, ChunkOverlap As Integer)

        If embed_store Is Nothing Then embed_store = New EmbeddingStore(EmbedModel, EmbedVocab)

        Dim doc = Application.ActiveDocument
        Dim docId = doc.FullName

        ' 0) Early return, wenn schon indexiert und kein Refresh gewünscht
        If embed_indexedDocs.Contains(docId) AndAlso Not refresh Then
            Return
        End If

        ' 1) Parameter validieren
        Dim nn As Integer = ChunkLength     ' Sätze pro Chunk
        Dim mm As Integer = ChunkOverlap     ' Überlappung
        Dim stepSize = nn - mm
        If nn <= 0 OrElse mm < 0 OrElse stepSize <= 0 Then
            Throw New ArgumentException("Bitte nn>0, mm≥0 und mm<nn wählen.")
        End If

        ' 2) Sätze holen und leere filtern
        Dim sentences = doc.Sentences.Cast(Of Range)() _
                        .Where(Function(r) Not String.IsNullOrWhiteSpace(r.Text)) _
                        .ToList()
        Dim total = sentences.Count

        If total < nn Then
            Return
        End If

        ' 3) Chunks bauen (nur volle nn-Satz-Chunks)
        Dim chunks As New List(Of TextChunk)()
        For idx As Integer = 0 To total - nn Step stepSize
            Dim startIdx = idx
            Dim endIdx = idx + nn - 1  ' garantiert ≤ total-1

            ' Text zusammenbauen
            Dim parts = sentences.Skip(startIdx).Take(nn).Select(Function(r) r.Text.Trim())
            Dim chunkText = String.Join(" ", parts)

            ' Sehr kurze Chunks überspringen
            If chunkText.Length < 10 Then
                Continue For
            End If

            ' 4) Offset berechnen – direkt aus Range.Start, kein doc.Range mehr
            Dim rangeStart = sentences(startIdx).Start
            Dim startOffset = If(rangeStart < 0, 0, rangeStart)
            Dim rangeEnd = sentences(endIdx).End

            ' Chunk hinzufügen
            chunks.Add(New TextChunk With {
            .Text = chunkText,
            .startOffset = startOffset,
            .EndOffset = rangeEnd
        })
        Next
        ' 4) Indexieren
        embed_store.IndexDocument(docId, chunks)
        If Not embed_indexedDocs.Contains(docId) Then embed_indexedDocs.Add(docId)
    End Sub

    Public Sub RunSearch_Embed(EmbedModel As String, EmbedVocab As String, SearchContext As String, DoNext As Boolean, DoRefresh As Boolean)

        Dim Prefix = "-CSE"

        Try

            ' 1) Parameter

            Dim ChunkLength As Integer = Default_Embed_Chunks
            Dim ChunkOverlap As Integer = Default_Embed_Overlap
            Dim Min_Score As Double = Default_Embed_Min_Score
            Dim Top_K As Integer = Default_Embed_Top_K
            Dim allDocs As Boolean = False
            Dim Fallback As Boolean = False


            If Not embed_indexedDocs.Contains(Application.ActiveDocument.FullName) Or DoRefresh Then

                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Sentences per chunk:", ChunkLength),
                    New SLib.InputParameter("Overlap per chunk", ChunkOverlap),
                    New SLib.InputParameter("Minimum relevance", Min_Score),
                    New SLib.InputParameter("Maximum hits", Top_K),
                    New SLib.InputParameter("Always hits", Fallback)
                    }

                If ShowCustomVariableInputForm("Please set your embedding and search values:", $"Context Search (Embedding)", params) Then

                    ChunkLength = CInt(params(0).Value)
                    ChunkOverlap = CInt(params(1).Value)
                    Min_Score = CDbl(params(2).Value)
                    Top_K = CInt(params(3).Value)
                    Fallback = CBool(params(4).Value)

                Else
                    Return
                End If

            Else
                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Minimum relevance", Min_Score),
                    New SLib.InputParameter("Maximum hits:", Top_K),
                    New SLib.InputParameter("Always hits", Fallback)
                    }
                If ShowCustomVariableInputForm("Please set your search values:", $"Context Search (Embedding)", params) Then

                    Min_Score = CDbl(params(0).Value)
                    Top_K = CInt(params(1).Value)
                    Fallback = CBool(params(2).Value)

                Else
                    Return
                End If

            End If

            ' 2) Für Next-Suche: Selektion zurücksetzen & Cursor ans Ende
            Dim selRange As Word.Range = Application.Selection.Range
            If DoNext Then selRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

            ' 3) Index ggf. neu aufbauen
            RunIndexing_Embed(DoRefresh, EmbedModel, EmbedVocab, ChunkLength, ChunkOverlap)

            Dim currentDocId = Application.ActiveDocument.FullName
            Dim cursorPos = selRange.Start

            If Not DoNext Then
                ' --- COMPLETE: Suche im (Rest)Dokument oder in allen Docs ---
                Dim rawHits = embed_store.Search(SearchContext, allDocs, True, currentDocId, cursorPos) _
                            .Where(Function(r) r.Score > 0 _
                                             AndAlso (allDocs OrElse r.StartOffset > cursorPos)) _
                            .OrderByDescending(Function(r) r.Score) _
                            .ToList()

                ' Treffer über Schwellwert
                Dim scoredHits = rawHits.Where(Function(r) r.Score >= Min_Score).ToList()
                Dim hits As List(Of SearchResult)
                If scoredHits.Count > 0 Then
                    hits = scoredHits.Take(Top_K).ToList()
                ElseIf Fallback Then
                    ' Fallback: die besten TOP_K unabhängig vom Score
                    hits = rawHits.Take(Top_K).ToList()
                End If

                If hits.Count = 0 Then
                    ShowCustomMessageBox($"No hits found for '{SearchContext}'" & If(Fallback, ".", " and minimum relevance of {Min_Score}."))
                    Return
                End If

                Dim trackChangesEnabled As Boolean = Application.ActiveDocument.TrackRevisions
                Application.ActiveDocument.TrackRevisions = True

                For Each r In hits
                    Dim doc = If(r.DocId = currentDocId,
                             Application.ActiveDocument,
                             Application.Documents.Open(r.DocId))
                    Dim rng = doc.Range(r.StartOffset, r.EndOffset)
                    doc.Comments.Add(rng, $"{AN5}{Prefix}: '{SearchContext}' (Score {r.Score:F3})")
                Next

                Application.ActiveDocument.TrackRevisions = trackChangesEnabled

                ShowCustomMessageBox($"{hits.Count} hits found for '{SearchContext}', a minimum relevance of {Min_Score} and a maximum of {Top_K} hits. Comments have been added to them.")
            Else
                ' --- NEXT: Suche nur ab Cursor im aktuellen Dokument ---
                Dim rawHits = embed_store.Search(SearchContext, False, True, currentDocId, cursorPos) _
                            .Where(Function(r) r.Score > 0 AndAlso r.StartOffset > cursorPos) _
                            .OrderByDescending(Function(r) r.Score) _
                            .ToList()

                ' Treffer über Schwellwert
                Dim scoredHits = rawHits.Where(Function(r) r.Score >= Min_Score).ToList()
                Dim hits As List(Of SearchResult)
                If scoredHits.Count > 0 Then
                    hits = scoredHits.Take(Top_K).ToList()
                ElseIf Fallback Then
                    ' Fallback: die besten TOP_K unabhängig vom Score
                    hits = rawHits.Take(Top_K).ToList()
                End If

                If hits.Count = 0 Then
                    ShowCustomMessageBox($"No (further) hits found for '{SearchContext}'" & If(Fallback, ".", " and minimum relevance of {Min_Score}."))
                    Return
                End If

                Dim trackChangesEnabled As Boolean = Application.ActiveDocument.TrackRevisions
                Application.ActiveDocument.TrackRevisions = True

                For Each r In hits
                    Dim doc = If(r.DocId = currentDocId,
                             Application.ActiveDocument,
                             Application.Documents.Open(r.DocId))
                    Dim rng = doc.Range(r.StartOffset, r.EndOffset)
                    doc.Comments.Add(rng, $"{AN5}{Prefix}: '{SearchContext}' (Score {r.Score:F3})")
                Next

                Application.ActiveDocument.TrackRevisions = trackChangesEnabled

                ShowCustomMessageBox($"The (next) {hits.Count} have been found for '{SearchContext}', with a maximum of {Top_K} hits. Comments have been added to them.")
            End If

        Catch ex As Exception
            MessageBox.Show("Error in RunSearch_Embed: " & ex.Message)
        End Try
    End Sub


    ' Bag-of-Words–Store und Index-Tracking
    Private store_bow As EmbeddingStore_BagofWords = New EmbeddingStore_BagofWords()
    Private indexedDocs_bow As HashSet(Of String) = New HashSet(Of String)()

    Public Sub RunIndexing_bow(refresh As Boolean, ChunkLength As Integer, ChunkOverlap As Integer)
        Dim doc = Application.ActiveDocument
        Dim docId = doc.FullName

        ' 0) Early return, wenn schon indexiert und kein Refresh gewünscht
        If indexedDocs_bow.Contains(docId) AndAlso Not refresh Then
            Return
        End If

        ' 1) Parameter validieren
        Dim nn As Integer = ChunkLength     ' Sätze pro Chunk
        Dim mm As Integer = ChunkOverlap     ' Überlappung
        Dim stepSize = nn - mm
        If nn <= 0 OrElse mm < 0 OrElse stepSize <= 0 Then
            Throw New System.ArgumentException("Bitte nn>0, mm≥0 und mm<nn wählen.")
        End If

        ' 2) Sätze holen und leere filtern
        Dim sentences = doc.Sentences.Cast(Of Word.Range)() _
                    .Where(Function(r) Not String.IsNullOrWhiteSpace(r.Text)) _
                    .ToList()
        Dim total = sentences.Count
        If total < nn Then
            Return
        End If

        ' 3) Chunks bauen (nur volle nn-Satz-Chunks)
        Dim chunks As New List(Of TextChunk)()
        For idx As Integer = 0 To total - nn Step stepSize
            Dim startIdx = idx
            Dim endIdx = idx + nn - 1  ' garantiert ≤ total-1

            ' Text zusammenbauen
            Dim parts = sentences.Skip(startIdx).Take(nn).Select(Function(r) r.Text.Trim())
            Dim chunkText = String.Join(" ", parts)
            If chunkText.Length < 10 Then Continue For

            ' Offsets aus den Ranges
            Dim rangeStart = sentences(startIdx).Start
            Dim startOffset = If(rangeStart < 0, 0, rangeStart)
            Dim rangeEnd = sentences(endIdx).End

            chunks.Add(New TextChunk With {
            .Text = chunkText,
            .startOffset = startOffset,
            .EndOffset = rangeEnd
        })
        Next

        ' 4) Indexieren
        store_bow.IndexDocument(docId, chunks)
        If Not indexedDocs_bow.Contains(docId) Then indexedDocs_bow.Add(docId)
    End Sub

    Public Sub RunSearch_bow(SearchContext As String, DoNext As Boolean, DoRefresh As Boolean)

        Dim Prefix = "-CSB"

        Try

            ' 1) Parameter

            Dim ChunkLength As Integer = Default_Embed_Chunks_bow
            Dim ChunkOverlap As Integer = Default_Embed_Overlap_bow
            Dim Min_Score As Double = Default_Embed_Min_Score
            Dim Top_K As Integer = Default_Embed_Top_K
            Dim allDocs As Boolean = False
            Dim Fallback As Boolean = False

            If embed_store Is Nothing Or DoRefresh Then

                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Sentences per chunk:", ChunkLength),
                    New SLib.InputParameter("Overlap per chunk", ChunkOverlap),
                    New SLib.InputParameter("Minimum relevance", Min_Score),
                    New SLib.InputParameter("Maximum hits", Top_K),
                    New SLib.InputParameter("Always hits", Fallback)
                    }

                If ShowCustomVariableInputForm("Please set your 'Bag of Words' and search values:", $"Context Search (Bag of Words)", params) Then

                    ChunkLength = CInt(params(0).Value)
                    ChunkOverlap = CInt(params(1).Value)
                    Min_Score = CDbl(params(2).Value)
                    Top_K = CInt(params(3).Value)
                    Fallback = CBool(params(4).Value)

                Else
                    Return
                End If

            Else
                Dim params() As SLib.InputParameter = {
                    New SLib.InputParameter("Minimum relevance", Min_Score),
                    New SLib.InputParameter("Maximum hits:", Top_K),
                    New SLib.InputParameter("Always hits", Fallback),
                    New SLib.InputParameter("Search all indexed docs:", allDocs)
                    }
                If ShowCustomVariableInputForm("Please set your search values:", $"Context Search (Bag of Words)", params) Then

                    Min_Score = CDbl(params(0).Value)
                    Top_K = CInt(params(1).Value)
                    Fallback = CBool(params(2).Value)
                    allDocs = CBool(params(3).Value)

                Else
                    Return
                End If

            End If

            ' 2) Für Next-Suche: Selektion zurücksetzen & Cursor ans Ende
            Dim selRange As Word.Range = Application.Selection.Range
            If DoNext Then selRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

            ' 3) Index ggf. neu aufbauen
            RunIndexing_bow(DoRefresh, ChunkLength, ChunkOverlap)

            Dim currentDocId = Application.ActiveDocument.FullName
            Dim cursorPos = selRange.Start

            If Not DoNext Then
                ' --- COMPLETE: Suche im (Rest-)Dokument oder in allen Docs ---
                Dim rawHits = store_bow.Search(SearchContext, allDocs, True, currentDocId, cursorPos) _
                        .Where(Function(r) r.Score > 0 _
                                         AndAlso (allDocs OrElse r.StartOffset > cursorPos)) _
                        .OrderByDescending(Function(r) r.Score) _
                        .ToList()

                ' Treffer über Schwellwert
                Dim scoredHits = rawHits.Where(Function(r) r.Score >= Min_Score).ToList()
                Dim hits As List(Of SearchResult)
                If scoredHits.Count > 0 Then
                    hits = scoredHits.Take(Top_K).ToList()
                ElseIf Fallback Then
                    ' Fallback: die besten TOP_K unabhängig vom Score
                    hits = rawHits.Take(Top_K).ToList()
                End If

                If hits.Count = 0 Then
                    ShowCustomMessageBox($"No hits found for '{SearchContext}'" & If(Fallback, ".", " And minimum relevance of {Min_Score}."))
                    Return
                End If

                Dim trackChangesEnabled As Boolean = Application.ActiveDocument.TrackRevisions
                Application.ActiveDocument.TrackRevisions = True

                For Each r In hits
                    Dim docTarget = If(r.DocId = currentDocId,
                                   Application.ActiveDocument,
                                   Application.Documents.Open(r.DocId))
                    Dim rng = docTarget.Range(r.StartOffset, r.EndOffset)
                    docTarget.Comments.Add(rng, $"{AN5}{Prefix}: '{SearchContext}' (BoW score {r.Score:F3})")
                Next

                Application.ActiveDocument.TrackRevisions = trackChangesEnabled

                ShowCustomMessageBox($"{hits.Count} hits found for '{SearchContext}', a minimum relevance of {Min_Score} and a maximum of {Top_K} hits. Comments have been added to them.")
            Else
                ' --- NEXT: Suche nur ab Cursor im aktuellen Dokument ---
                Dim rawHits = store_bow.Search(SearchContext, False, True, currentDocId, cursorPos) _
                        .Where(Function(r) r.Score > 0 AndAlso r.StartOffset > cursorPos) _
                        .OrderByDescending(Function(r) r.Score) _
                        .ToList()

                Dim scoredHits = rawHits.Where(Function(r) r.Score >= Min_Score).ToList()
                Dim hits As List(Of SearchResult)
                If scoredHits.Count > 0 Then
                    hits = scoredHits.Take(Top_K).ToList()
                ElseIf Fallback Then
                    ' Fallback: die besten TOP_K unabhängig vom Score
                    hits = rawHits.Take(Top_K).ToList()
                End If

                If hits.Count = 0 Then
                    ShowCustomMessageBox($"No (further) hits found for '{SearchContext}'" & If(Fallback, ".", " And minimum relevance of {Min_Score}."))
                    Return
                End If

                Dim trackChangesEnabled As Boolean = Application.ActiveDocument.TrackRevisions
                Application.ActiveDocument.TrackRevisions = True

                For Each r In hits
                    Dim docTarget = If(r.DocId = currentDocId,
                                   Application.ActiveDocument,
                                   Application.Documents.Open(r.DocId))
                    Dim rng = docTarget.Range(r.StartOffset, r.EndOffset)
                    docTarget.Comments.Add(rng, $"{AN5}{Prefix}: '{SearchContext}' (BoW score {r.Score:F3})")
                Next

                Application.ActiveDocument.TrackRevisions = trackChangesEnabled

                ShowCustomMessageBox($"The (next) {hits.Count} have been found for '{SearchContext}', with a maximum of {Top_K} hits. Comments have been added to them.")
            End If
        Catch ex As System.Exception
            MessageBox.Show("Error in RunSearch_BoW: " & ex.Message)
        End Try
    End Sub


End Class
