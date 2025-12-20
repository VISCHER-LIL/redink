' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Slib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' Shared CSS styling for HTML summary windows
    Private Const SummaryHtmlStyle As String =
        "<style>" &
        "body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 10pt; line-height: 1.5; padding: 20px; margin: 0; }" &
        "ul, ol { margin-left: 20px; }" &
        "li { margin-bottom: 6px; }" &
        "h1, h2, h3 { color: #333; }" &
        "strong { color: #003366; }" &
        "code { background: #f6f8fa; padding: 2px 4px; border-radius: 3px; }" &
        "pre { background: #f6f8fa; padding: 10px; border-radius: 4px; overflow-x: auto; }" &
        "</style>"

    ' Compares the currently active Word document with another open Word document,
    ' exports the comparison to filtered HTML, and shows it via ShowHTMLCustomMessageBox().
    Public Shared Sub CompareActiveDocWithOtherOpenDoc()
        ' Acquire the running Word instance
        Dim wordAppObj As Object = Nothing
        Try
            wordAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
        Catch
            ShowCustomMessageBox("Microsoft Word is not running or cannot be accessed.", AN)
            Exit Sub
        End Try

        Dim wordApp As Microsoft.Office.Interop.Word.Application = TryCast(wordAppObj, Microsoft.Office.Interop.Word.Application)
        If wordApp Is Nothing Then
            ShowCustomMessageBox("Unable to access the Word application.", AN)
            Exit Sub
        End If

        ' Ensure there is an active document
        If wordApp.Documents Is Nothing OrElse wordApp.Documents.Count = 0 Then
            ShowCustomMessageBox("No document is open in Word.", AN)
            Exit Sub
        End If

        Dim activeDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Try
            activeDoc = wordApp.ActiveDocument
        Catch
        End Try
        If activeDoc Is Nothing Then
            ShowCustomMessageBox("No active document detected in Word.", AN)
            Exit Sub
        End If

        ' Build the list of other open documents
        Dim otherDocs As New List(Of Microsoft.Office.Interop.Word.Document)()
        For Each d As Microsoft.Office.Interop.Word.Document In wordApp.Documents
            If Not Object.ReferenceEquals(d, activeDoc) Then
                otherDocs.Add(d)
            End If
        Next

        If otherDocs.Count = 0 Then
            ShowCustomMessageBox("No other open document found to compare against.", AN)
            Exit Sub
        End If

        ' Pick the second document (auto if only one; otherwise ask via SLib.SelectValue)
        Dim docToCompare As Microsoft.Office.Interop.Word.Document = Nothing
        If otherDocs.Count = 1 Then
            docToCompare = otherDocs(0)
        Else
            Dim items As New List(Of Slib.SelectionItem)()
            Dim indexToDoc As New Dictionary(Of Integer, Microsoft.Office.Interop.Word.Document)()
            Dim idx As Integer = 1
            For Each d In otherDocs
                Dim disp As String
                Try
                    disp = If(String.IsNullOrEmpty(d.Name), "(unnamed document)", d.Name)
                Catch
                    disp = "(document)"
                End Try
                items.Add(New Slib.SelectionItem(disp, idx))
                indexToDoc(idx) = d
                idx += 1
            Next

            Dim chosenIdx As Integer = Slib.SelectValue(items, 1, "Select the document to compare with:", $"{AN} Compare")
            If chosenIdx <= 0 OrElse Not indexToDoc.ContainsKey(chosenIdx) Then Exit Sub
            docToCompare = indexToDoc(chosenIdx)
        End If

        ' Compare and export -> filtered HTML
        Dim compareDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim tempHtmlPath As String = Nothing
        Dim tempFolder As String = Nothing

        ' UI suppression to reduce flicker
        Dim prevScreenUpdating As Boolean = wordApp.ScreenUpdating
        Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel = wordApp.DisplayAlerts
        Dim prevWindow As Microsoft.Office.Interop.Word.Window = Nothing

        ' Store extracted changes for LLM summarization
        Dim extractedChangesText As String = Nothing

        Try
            wordApp.ScreenUpdating = False
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
            prevWindow = wordApp.ActiveWindow

            ' Create comparison (returns a Document)
            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=activeDoc,
                RevisedDocument:=docToCompare,
                Destination:=WdCompareDestination.wdCompareDestinationNew,
                Granularity:=WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=True,
                CompareCaseChanges:=True,
                CompareWhitespace:=True,
                CompareTables:=True,
                CompareHeaders:=True,
                CompareFootnotes:=True,
                CompareTextboxes:=True,
                CompareFields:=True,
                CompareComments:=True,
                CompareMoves:=True,
                RevisedAuthor:=Environment.UserName,
                IgnoreAllComparisonWarnings:=False
            )
            If compareDoc Is Nothing Then
                ShowCustomMessageBox("Word did not produce a comparison document.", AN)
                Exit Sub
            End If

            ' Keep its window hidden (best effort)
            Try
                If compareDoc.Windows IsNot Nothing AndAlso compareDoc.Windows.Count > 0 Then
                    compareDoc.Windows(1).Visible = False
                End If
            Catch
            End Try

            ' Extract changes with markup tags for LLM summarization
            extractedChangesText = ExtractChangesWithMarkupTags(compareDoc)

            ' Export to filtered HTML
            tempFolder = Path.Combine(Path.GetTempPath(), $"{AN2}_compare_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempFolder)
            tempHtmlPath = Path.Combine(tempFolder, "comparison.htm")

            compareDoc.SaveAs2(FileName:=tempHtmlPath, FileFormat:=WdSaveFormat.wdFormatFilteredHTML)

            ' Restore focus ASAP to reduce flicker
            Try
                prevWindow?.Activate()
            Catch
            End Try

            ' Close the comparison doc to release file locks
            Try
                compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            Catch
            End Try
            compareDoc = Nothing

            ' Read bytes with retry to avoid transient locks
            Dim raw As Byte() = Nothing
            Dim maxAttempts As Integer = 10
            Dim delayMs As Integer = 100
            For attempt As Integer = 1 To maxAttempts
                Try
                    raw = File.ReadAllBytes(tempHtmlPath)
                    Exit For
                Catch ex As IOException
                    Threading.Thread.Sleep(delayMs)
                End Try
            Next
            If raw Is Nothing OrElse raw.Length = 0 Then
                ShowCustomMessageBox($"Comparison failed: could not read '{tempHtmlPath}'.", AN)
                Exit Sub
            End If

            ' Decode using BOM or <meta charset>, else default to Windows-1252
            Dim enc As System.Text.Encoding = System.Text.Encoding.UTF8
            If raw.Length >= 3 AndAlso raw(0) = &HEF AndAlso raw(1) = &HBB AndAlso raw(2) = &HBF Then
                enc = System.Text.Encoding.UTF8
            Else
                Dim probe As String = System.Text.Encoding.GetEncoding(28591).GetString(raw) ' ISO-8859-1
                Dim m As System.Text.RegularExpressions.Match =
                    System.Text.RegularExpressions.Regex.Match(
                        probe,
                        "(?is)<meta[^>]*?(?:charset\s*=\s*[""']?\s*([A-Za-z0-9_\-]+)|http-equiv\s*=\s*[""']?\s*content-type[""'][^>]*?content\s*=\s*[""'][^""']*?;\s*charset\s*=\s*([A-Za-z0-9_\-]+))",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                Dim charset As String = Nothing
                If m.Success Then
                    charset = If(m.Groups(1).Success, m.Groups(1).Value, If(m.Groups(2).Success, m.Groups(2).Value, Nothing))
                End If
                If Not String.IsNullOrEmpty(charset) Then
                    Try
                        enc = System.Text.Encoding.GetEncoding(charset)
                    Catch
                        enc = System.Text.Encoding.GetEncoding(1252)
                    End Try
                Else
                    enc = System.Text.Encoding.GetEncoding(1252)
                End If
            End If

            Dim html As String = enc.GetString(raw)

            ' Ensure proper meta charset and inject base for resources
            Dim hasHead As Boolean = html.IndexOf("<head", StringComparison.OrdinalIgnoreCase) >= 0
            Dim metaCharset As String = $"<meta http-equiv=""Content-Type"" content=""text/html; charset={enc.WebName}"">"
            If hasHead Then
                Dim rxHead As New System.Text.RegularExpressions.Regex("(<head[^>]*>)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                html = rxHead.Replace(html, "$1" & metaCharset, 1)
            Else
                html = "<html><head>" & metaCharset & "</head>" & html.Replace("<html>", "").Replace("</html>", "") & "</html>"
            End If

            Dim baseTag As String = $"<base href=""{tempFolder.Replace("\", "/")}/"">"
            Dim rxMetaEnd As New System.Text.RegularExpressions.Regex("(<head[^>]*>)(.*?)(</head>)", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline)
            html = rxMetaEnd.Replace(html,
                                     Function(mm)
                                         Return mm.Groups(1).Value & baseTag & mm.Groups(2).Value & mm.Groups(3).Value
                                     End Function, 1)

            ' Show result with optional "Summarize Changes" button
            Dim extraAction As System.Action = Nothing
            If Not String.IsNullOrWhiteSpace(extractedChangesText) Then
                extraAction = Sub()
                                  ' This runs on the STA thread of ShowHTMLCustomMessageBox
                                  SummarizeComparisonChangesAsync(extractedChangesText)
                              End Sub
            End If

            ShowHTMLCustomMessageBox(html, $"{AN} Word Active Compare",
                                     extraButtonText:=If(Not String.IsNullOrWhiteSpace(extractedChangesText), "Summarize Changes", Nothing),
                                     extraButtonAction:=extraAction,
                                     CloseAfterExtra:=False)

        Catch ex As System.Exception
            ShowCustomMessageBox($"Comparison failed: {ex.Message}", AN)
        Finally
            ' Safety close
            If compareDoc IsNot Nothing Then
                Try
                    compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                Catch
                End Try
                compareDoc = Nothing
            End If

            ' Restore UI
            wordApp.DisplayAlerts = prevAlerts
            wordApp.ScreenUpdating = prevScreenUpdating

            ' Cleanup temp (delayed, best effort)
            If Not String.IsNullOrEmpty(tempFolder) Then
                Try
                    Dim t As New Threading.Thread(
                        Sub()
                            Try
                                Threading.Thread.Sleep(3000)
                                If File.Exists(tempHtmlPath) Then File.Delete(tempHtmlPath)
                                Dim filesFolder As String = tempHtmlPath & "_files"
                                If Directory.Exists(filesFolder) Then
                                    Try
                                        Directory.Delete(filesFolder, recursive:=True)
                                    Catch
                                    End Try
                                End If
                                Directory.Delete(tempFolder, recursive:=True)
                            Catch
                            End Try
                        End Sub)
                    t.IsBackground = True
                    t.Start()
                Catch
                End Try
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Extracts text from a comparison document with revisions marked using &lt;ins&gt; and &lt;del&gt; tags,
    ''' comments marked with &lt;comment&gt; tags, and footnotes/endnotes included.
    ''' </summary>
    Private Shared Function ExtractChangesWithMarkupTags(compareDoc As Microsoft.Office.Interop.Word.Document) As String
        If compareDoc Is Nothing Then Return String.Empty

        Dim sb As New System.Text.StringBuilder()

        Try
            ' Process main document content with revisions
            Dim content As Microsoft.Office.Interop.Word.Range = compareDoc.Content

            ' Build text with revision markup
            For Each para As Microsoft.Office.Interop.Word.Paragraph In compareDoc.Paragraphs
                Dim paraText As New System.Text.StringBuilder()
                Dim rng As Microsoft.Office.Interop.Word.Range = para.Range

                ' Check each character/word for revisions
                For Each rev As Microsoft.Office.Interop.Word.Revision In rng.Revisions
                    Try
                        Dim revText As String = If(rev.Range.Text, String.Empty)
                        If String.IsNullOrEmpty(revText) Then Continue For

                        Select Case rev.Type
                            Case WdRevisionType.wdRevisionInsert
                                paraText.Append($"<ins>{revText}</ins>")
                            Case WdRevisionType.wdRevisionDelete
                                paraText.Append($"<del>{revText}</del>")
                            Case WdRevisionType.wdRevisionMovedFrom
                                paraText.Append($"<del>[moved from:]{revText}</del>")
                            Case WdRevisionType.wdRevisionMovedTo
                                paraText.Append($"<ins>[moved to:]{revText}</ins>")
                            Case Else
                                ' For formatting and other changes, note them but include the text
                                paraText.Append($"<ins>[{rev.Type}:]{revText}</ins>")
                        End Select
                    Catch
                    End Try
                Next

                ' If no revisions in this paragraph, just add the plain text
                If paraText.Length = 0 Then
                    Try
                        paraText.Append(If(rng.Text, String.Empty))
                    Catch
                    End Try
                End If

                sb.AppendLine(paraText.ToString())
            Next

            ' Process comments
            If compareDoc.Comments IsNot Nothing AndAlso compareDoc.Comments.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("<comments>")
                For Each cmt As Microsoft.Office.Interop.Word.Comment In compareDoc.Comments
                    Try
                        Dim author As String = If(cmt.Author, "Unknown")
                        Dim commentText As String = If(cmt.Range.Text, String.Empty)
                        Dim scopeText As String = String.Empty
                        Try
                            scopeText = If(cmt.Scope.Text, String.Empty)
                        Catch
                        End Try

                        sb.AppendLine($"<comment author=""{System.Security.SecurityElement.Escape(author)}"" scope=""{System.Security.SecurityElement.Escape(scopeText)}"">{System.Security.SecurityElement.Escape(commentText)}</comment>")
                    Catch
                    End Try
                Next
                sb.AppendLine("</comments>")
            End If

            ' Process footnotes
            If compareDoc.Footnotes IsNot Nothing AndAlso compareDoc.Footnotes.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("<footnotes>")
                For Each fn As Microsoft.Office.Interop.Word.Footnote In compareDoc.Footnotes
                    Try
                        Dim fnText As String = If(fn.Range.Text, String.Empty)
                        sb.AppendLine($"<footnote index=""{fn.Index}"">{System.Security.SecurityElement.Escape(fnText)}</footnote>")
                    Catch
                    End Try
                Next
                sb.AppendLine("</footnotes>")
            End If

            ' Process endnotes
            If compareDoc.Endnotes IsNot Nothing AndAlso compareDoc.Endnotes.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("<endnotes>")
                For Each en As Microsoft.Office.Interop.Word.Endnote In compareDoc.Endnotes
                    Try
                        Dim enText As String = If(en.Range.Text, String.Empty)
                        sb.AppendLine($"<endnote index=""{en.Index}"">{System.Security.SecurityElement.Escape(enText)}</endnote>")
                    Catch
                    End Try
                Next
                sb.AppendLine("</endnotes>")
            End If

        Catch ex As System.Exception
            sb.AppendLine($"[Error extracting changes: {ex.Message}]")
        End Try

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Calls the LLM to summarize the extracted changes and displays the result in a new HTML window.
    ''' This method spawns its own async operation and ShowHTMLCustomMessageBox call.
    ''' </summary>
    Private Shared Sub SummarizeComparisonChangesAsync(extractedChangesText As String)
        ' Run the LLM call and display on a new thread to avoid blocking
        Dim t As New Threading.Thread(
            Sub()
                Try
                    ' Build the prompt
                    Dim userPrompt As String = "<TEXTTOPROCESS>" & vbCrLf & extractedChangesText & vbCrLf & "</TEXTTOPROCESS>"

                    ' System prompt for change analysis
                    Dim systemPrompt As String = SP_Markup

                    Dim llmResult As String = String.Empty
                    Try
                        llmResult = SharedMethods.LLM(
                            _context,
                            systemPrompt,
                            userPrompt,
                            "",
                            "",
                            0,
                            False,
                            False).GetAwaiter().GetResult()
                    Catch ex As System.Exception
                        llmResult = $"Error calling LLM: {ex.Message}"
                    End Try

                    ' Convert Markdown to HTML using Markdig
                    Dim htmlResult As String
                    Try
                        Dim pipeline = New Markdig.MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
                        Dim bodyHtml As String = Markdig.Markdown.ToHtml(If(llmResult, String.Empty), pipeline)

                        htmlResult = "<!DOCTYPE html><html><head><meta charset=""utf-8"">" &
                                     SummaryHtmlStyle &
                                     "</head><body>" &
                                     bodyHtml &
                                     "</body></html>"
                    Catch ex As System.Exception
                        htmlResult = $"<html><body><pre>{System.Security.SecurityElement.Escape(If(llmResult, ex.Message))}</pre></body></html>"
                    End Try

                    ShowHTMLCustomMessageBox(htmlResult, $"{AN} Change Summary")

                Catch ex As System.Exception
                    ShowCustomMessageBox($"Failed to summarize changes: {ex.Message}", AN)
                End Try
            End Sub)
        t.SetApartmentState(Threading.ApartmentState.STA)
        t.IsBackground = True
        t.Start()
    End Sub


    ''' <summary>
    ''' Extracts revisions and comments from the active document or selection based on a date filter,
    ''' then summarizes them using the LLM and displays the result.
    ''' </summary>
    Public Shared Async Sub SummarizeDocumentChanges()
        Try
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing

            Try
                app = Globals.ThisAddIn.Application
                doc = app.ActiveDocument
            Catch
                ShowCustomMessageBox("No active document found.", AN)
                Exit Sub
            End Try

            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.", AN)
                Exit Sub
            End If

            ' Determine scope: selection or entire document
            Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection
            Dim useEntireDoc As Boolean = (sel Is Nothing OrElse sel.Range Is Nothing OrElse sel.Start = sel.End)
            Dim scopeRange As Microsoft.Office.Interop.Word.Range = If(useEntireDoc, doc.Content, sel.Range)
            Dim scopeDescription As String = If(useEntireDoc, "the entire document", "the selected text")

            ' Prompt for date filter
            Dim defaultDate As String = System.DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
            Dim userDateInput As String = ShowCustomInputBox(
                $"Enter the earliest date for changes to include (leave empty to include all tracked changes and only comments made not older than 60 minutes before the first change).{vbCrLf}{vbCrLf}Changes from {scopeDescription} will be analyzed.",
                $"{AN} Summarize Changes",
                True,
                defaultDate)

            If userDateInput Is Nothing Then
                ' User cancelled
                Exit Sub
            End If

            userDateInput = userDateInput.Trim()

            Dim filterDate As System.DateTime? = Nothing
            Dim filterByDate As Boolean = False

            If Not String.IsNullOrEmpty(userDateInput) Then
                Dim parsed As System.DateTime
                If System.DateTime.TryParse(userDateInput, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, parsed) Then
                    filterDate = parsed.Date ' Use start of day
                    filterByDate = True
                Else
                    ShowCustomMessageBox("Invalid date format. Operation aborted.", AN)
                    Exit Sub
                End If
            End If

            ' Extract revisions and comments
            Dim extractedText As String = ExtractRevisionsAndCommentsWithMarkup(doc, scopeRange, filterDate, filterByDate)

            If String.IsNullOrWhiteSpace(extractedText) Then
                Dim dateInfo As String = If(filterByDate, $" on or after {filterDate.Value:yyyy-MM-dd}", "")
                ShowCustomMessageBox($"No revisions or comments found{dateInfo} in {scopeDescription}.", AN)
                Exit Sub
            End If

            ' Build the prompt
            Dim userPrompt As String = "<TEXTTOPROCESS>" & vbCrLf & extractedText & vbCrLf & "</TEXTTOPROCESS>"

            ' System prompt for change analysis (same as SummarizeComparisonChangesAsync)
            Dim systemPrompt As String = SP_Markup

            Dim llmResult As String = String.Empty
            Try
                llmResult = Await SharedMethods.LLM(
                    _context,
                    systemPrompt,
                    userPrompt,
                    "",
                    "",
                    0,
                    False,
                    False)
            Catch ex As System.Exception
                llmResult = $"Error calling LLM: {ex.Message}"
            End Try

            ' Convert Markdown to HTML using Markdig
            Dim htmlResult As String
            Try
                Dim pipeline = New Markdig.MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
                Dim bodyHtml As String = Markdig.Markdown.ToHtml(If(llmResult, String.Empty), pipeline)

                Dim dateFilterInfo As String = If(filterByDate,
                    $"<p style='color:#666; font-size:9pt;'>Covering changes/comments from {filterDate.Value:yyyy-MM-dd} onwards in {scopeDescription}</p>",
                    $"<p style='color:#666; font-size:9pt;'>Covering all tracked changes in {scopeDescription} (and comments not older than 60 minutes before the first change)</p>")

                htmlResult = "<!DOCTYPE html><html><head><meta charset=""utf-8"">" &
                             SummaryHtmlStyle &
                             "</head><body>" &
                             dateFilterInfo &
                             bodyHtml &
                             "</body></html>"
            Catch ex As System.Exception
                htmlResult = $"<html><body><pre>{System.Security.SecurityElement.Escape(If(llmResult, ex.Message))}</pre></body></html>"
            End Try

            ' Show the result
            ShowHTMLCustomMessageBox(htmlResult, $"{AN} Change Summary")

        Catch ex As System.Exception
            ShowCustomMessageBox($"Failed to summarize changes: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Extracts revisions and comments from a document range with markup tags.
    ''' Uses the same format as ExtractChangesWithMarkupTags for LLM compatibility.
    ''' Ignores pure formatting revisions for output but uses them for comment date calculation.
    ''' If filterByDate is True, includes revisions and comments on or after filterDate.
    ''' If filterByDate is False, includes all substantive revisions and comments added since first revision minus 60 minutes.
    ''' </summary>
    Private Shared Function ExtractRevisionsAndCommentsWithMarkup(
        doc As Microsoft.Office.Interop.Word.Document,
        scopeRange As Microsoft.Office.Interop.Word.Range,
        filterDate As System.DateTime?,
        filterByDate As Boolean) As String

        If doc Is Nothing OrElse scopeRange Is Nothing Then Return String.Empty

        Dim sb As New System.Text.StringBuilder()
        Dim hasContent As Boolean = False

        ' Revision types that are pure formatting (to be ignored in output, but used for date calculation)
        Dim formattingTypes As New HashSet(Of Integer)({
            CInt(WdRevisionType.wdRevisionProperty),
            CInt(WdRevisionType.wdRevisionParagraphNumber),
            CInt(WdRevisionType.wdRevisionParagraphProperty),
            CInt(WdRevisionType.wdRevisionSectionProperty),
            CInt(WdRevisionType.wdRevisionStyle),
            CInt(WdRevisionType.wdRevisionStyleDefinition),
            CInt(WdRevisionType.wdRevisionTableProperty)
        })

        Try
            ' Find the earliest revision date (including formatting revisions) for comment filtering
            Dim earliestRevisionDate As System.DateTime? = Nothing
            For Each rev As Microsoft.Office.Interop.Word.Revision In scopeRange.Revisions
                Try
                    If Not earliestRevisionDate.HasValue OrElse rev.Date < earliestRevisionDate.Value Then
                        earliestRevisionDate = rev.Date
                    End If
                Catch
                End Try
            Next

            ' Calculate comment cutoff date: earliest revision minus 60 minutes
            Dim commentCutoffDate As System.DateTime? = Nothing
            If earliestRevisionDate.HasValue Then
                commentCutoffDate = earliestRevisionDate.Value.AddMinutes(-60)
            End If

            ' Collect substantive revisions within the scope range
            Dim revisionList As New List(Of Microsoft.Office.Interop.Word.Revision)()

            For Each rev As Microsoft.Office.Interop.Word.Revision In scopeRange.Revisions
                Try
                    ' Skip pure formatting revisions for output
                    If formattingTypes.Contains(CInt(rev.Type)) Then
                        Continue For
                    End If

                    Dim includeRevision As Boolean = False

                    If filterByDate Then
                        ' Include if revision date >= filter date
                        If rev.Date >= filterDate.Value Then
                            includeRevision = True
                        End If
                    Else
                        ' No date filter: include all substantive revisions
                        includeRevision = True
                    End If

                    If includeRevision Then
                        revisionList.Add(rev)
                    End If
                Catch
                End Try
            Next

            ' Sort revisions by position in document
            revisionList.Sort(Function(a, b)
                                  Try
                                      Return a.Range.Start.CompareTo(b.Range.Start)
                                  Catch
                                      Return 0
                                  End Try
                              End Function)

            ' Build revision output using same format as ExtractChangesWithMarkupTags
            For Each rev In revisionList
                Try
                    Dim revText As String = If(rev.Range.Text, String.Empty)
                    If String.IsNullOrEmpty(revText) Then Continue For

                    Select Case rev.Type
                        Case WdRevisionType.wdRevisionInsert
                            sb.AppendLine($"<ins>{revText}</ins>")
                            hasContent = True
                        Case WdRevisionType.wdRevisionDelete
                            sb.AppendLine($"<del>{revText}</del>")
                            hasContent = True
                        Case WdRevisionType.wdRevisionMovedFrom
                            sb.AppendLine($"<del>[moved from:]{revText}</del>")
                            hasContent = True
                        Case WdRevisionType.wdRevisionMovedTo
                            sb.AppendLine($"<ins>[moved to:]{revText}</ins>")
                            hasContent = True
                        Case Else
                            ' Other non-formatting revision types
                            sb.AppendLine($"<ins>[{rev.Type}:]{revText}</ins>")
                            hasContent = True
                    End Select
                Catch
                End Try
            Next

            ' Collect comments within the scope range
            Dim commentList As New List(Of Microsoft.Office.Interop.Word.Comment)()

            For Each cmt As Microsoft.Office.Interop.Word.Comment In doc.Comments
                Try
                    ' Check if comment scope overlaps with our range
                    Dim cmtStart As Integer = -1
                    Dim cmtEnd As Integer = -1
                    Try
                        cmtStart = cmt.Scope.Start
                        cmtEnd = cmt.Scope.End
                    Catch
                        Try
                            cmtStart = cmt.Reference.Start
                            cmtEnd = cmt.Reference.End
                        Catch
                            Continue For
                        End Try
                    End Try

                    ' Check if comment is within scope range
                    If cmtStart >= scopeRange.Start AndAlso cmtEnd <= scopeRange.End Then
                        Dim includeComment As Boolean = False

                        If filterByDate Then
                            ' Include if comment date >= filter date
                            If cmt.Date >= filterDate.Value Then
                                includeComment = True
                            End If
                        Else
                            ' No date filter: include comments added since first revision minus 60 minutes
                            If commentCutoffDate.HasValue Then
                                If cmt.Date >= commentCutoffDate.Value Then
                                    includeComment = True
                                End If
                            Else
                                ' No revisions found, so no comments to include
                                includeComment = False
                            End If
                        End If

                        If includeComment Then
                            commentList.Add(cmt)
                        End If
                    End If
                Catch
                End Try
            Next

            ' Sort comments by position
            commentList.Sort(Function(a, b)
                                 Try
                                     Return a.Scope.Start.CompareTo(b.Scope.Start)
                                 Catch
                                     Return 0
                                 End Try
                             End Function)

            ' Build comments output using same format as ExtractChangesWithMarkupTags
            If commentList.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("<comments>")
                For Each cmt In commentList
                    Try
                        Dim author As String = If(cmt.Author, "Unknown")
                        Dim commentText As String = If(cmt.Range.Text, String.Empty)
                        Dim scopeText As String = String.Empty
                        Try
                            scopeText = If(cmt.Scope.Text, String.Empty)
                        Catch
                        End Try

                        sb.AppendLine($"<comment author=""{System.Security.SecurityElement.Escape(author)}"" scope=""{System.Security.SecurityElement.Escape(scopeText)}"">{System.Security.SecurityElement.Escape(commentText)}</comment>")
                        hasContent = True
                    Catch
                    End Try
                Next
                sb.AppendLine("</comments>")
            End If

        Catch ex As System.Exception
            sb.AppendLine($"[Error extracting changes: {ex.Message}]")
        End Try

        Return If(hasContent, sb.ToString(), String.Empty)
    End Function

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
