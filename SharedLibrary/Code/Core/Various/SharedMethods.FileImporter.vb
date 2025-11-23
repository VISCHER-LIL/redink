' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public Shared Function ReadTextFile(filePath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                ' Normalize and check the path
                filePath = Path.GetFullPath(filePath)
                If Not File.Exists(filePath) Then
                    Return If(ReturnErrorInsteadOfEmpty, "Error: File not found.", "")
                End If

                ' Use StreamReader for reading
                Using reader As New StreamReader(filePath, System.Text.Encoding.UTF8, True)
                    Dim content As String = reader.ReadToEnd()
                    Return content
                End Using
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading file: {ex.Message}", "")
            End Try
        End Function

        Public Shared Function ReadRtfAsText(ByVal rtfPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Try
                Dim rtfContent As String = File.ReadAllText(rtfPath)
                Using rtb As New RichTextBox()
                    rtb.Visible = False
                    rtb.Rtf = rtfContent
                    Return rtb.Text
                End Using
            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading RTF: {ex.Message}", "")
            End Try
        End Function

        Public Shared Function ReadWordDocument(ByVal docPath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
            Dim app As Microsoft.Office.Interop.Word.Application = Nothing
            Dim doc As Document = Nothing

            Try
                Try
                    ' Try to attach to an existing Word instance.
                    app = CType(Marshal.GetActiveObject("Word.Application"), Microsoft.Office.Interop.Word.Application)
                Catch ex As System.Exception
                    ' If Word is not running, create a new Word application.
                    app = New Microsoft.Office.Interop.Word.Application With {.Visible = False}
                End Try

                ' Open the Word document in read-only mode                
                Dim fileName As Object = docPath
                doc = app.Documents.Open(fileName, [ReadOnly]:=True, Visible:=False)

                ' Extract the content text
                Dim text As String = doc.Content.Text

                ' Close the document without saving changes
                doc.Close(SaveChanges:=False)

                ' Return the extracted text
                Return text

            Catch ex As System.Exception
                ' Ensure the document is closed in case of an error
                If doc IsNot Nothing Then
                    doc.Close(SaveChanges:=False)
                End If

                ' Return the error message (or empty string if ReturnErrorInsteadOfEmpty=False)
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading Word document: {ex.Message}", "")

            Finally
                ' Only quit the application if it was newly created
                If app IsNot Nothing AndAlso app.Visible = False Then
                    app.Quit()
                End If
            End Try
        End Function

        Public Shared Async Function ReadPdfAsText(ByVal pdfPath As String,
                                            Optional ByVal ReturnErrorInsteadOfEmpty As Boolean = True,
                                            Optional ByVal DoOCR As Boolean = False,
                                            Optional ByVal AskUser As Boolean = True,
                                            Optional ByVal context As ISharedContext = Nothing) As Task(Of String)
            Try
                Dim sb As New System.Text.StringBuilder()

                If Not System.IO.File.Exists(pdfPath) Then
                    Throw New System.IO.FileNotFoundException("PDF not found.", pdfPath)
                End If

                Dim pageCount As Integer = 0
                Dim totalLetters As Integer = 0

                ' Per-page diagnostics (evaluated only if DoOCR=True)
                Dim pagesImageOnly As New System.Collections.Generic.List(Of Integer)()
                Dim pagesLowLetters As New System.Collections.Generic.List(Of Integer)()
                Dim minLettersThreshold As Integer = 15 ' per-page "few letters" threshold

                ' Open the PDF document
                Dim parseOptions As New UglyToad.PdfPig.ParsingOptions() With {
                        .UseLenientParsing = True
                    }

                Using document As UglyToad.PdfPig.PdfDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath, parseOptions)
                    Dim pageTotal As Integer = document.NumberOfPages
                    pageCount = pageTotal

                    Debug.WriteLine("PDF has " & pageTotal & " pages.")

                    For pageNumber As Integer = 1 To pageTotal
                        Dim page As UglyToad.PdfPig.Content.Page = Nothing

                        Try
                            page = document.GetPage(pageNumber)
                        Catch ex As Exception
                            Debug.WriteLine($"PDF page {pageNumber} failed To load: {ex.Message}")
                            Continue For
                        End Try

                        Dim pageText As String = Nothing

                        ' Extract text from the page safely. If this fails, try a very simple fallback.
                        Try
                            pageText = ExtractPageTextFromPdf(page)
                            If Not String.IsNullOrWhiteSpace(pageText) Then
                                sb.AppendLine(pageText)
                            End If
                            Debug.WriteLine("Page " & pageNumber & " extracted text length: " & If(pageText IsNot Nothing, pageText.Length, 0))
                        Catch ex As Exception
                            ' Last-resort fallback: concatenate raw letters so we don't lose the page entirely.
                            Try
                                Dim letters = page.Letters
                                If letters IsNot Nothing AndAlso letters.Count > 0 Then
                                    Dim raw As New System.Text.StringBuilder(letters.Count)
                                    For Each l In letters
                                        raw.Append(l.Value)
                                    Next
                                    sb.AppendLine(raw.ToString())
                                End If
                            Catch
                                ' Give up on this page; continue with the rest.
                            End Try
                        End Try

                        ' Count letters and detect image-only/few-letters pages for OCR heuristic (only if DoOCR=True)
                        If DoOCR Then
                            Dim lettersThis As Integer = 0
                            Try
                                lettersThis = page.Letters.Count()
                                totalLetters += lettersThis
                            Catch
                                ' ignore
                            End Try

                            ' Page-level triggers
                            If String.IsNullOrWhiteSpace(pageText) AndAlso lettersThis = 0 Then
                                pagesImageOnly.Add(pageNumber)
                            ElseIf lettersThis < minLettersThreshold Then
                                pagesLowLetters.Add(pageNumber)
                            End If
                        End If
                    Next
                End Using

                Dim extractedText As String = sb.ToString()

                ' Disable OCR if no OCR-capable call is configured or context missing
                If DoOCR AndAlso (context Is Nothing OrElse System.String.IsNullOrWhiteSpace(context.INI_APICall_Object)) Then
                    DoOCR = False
                End If

                ' If DoOCR is disabled → just return whatever text we found (or empty string)
                If Not DoOCR Then
                    Return extractedText
                End If

                ' --- Heuristics only evaluated when DoOCR=True ---
                Dim shouldSuggestOcr As Boolean = False
                Dim reasons As New System.Collections.Generic.List(Of String)()

                ' Gather metrics for heuristic evaluation
                Dim fileLen As Long = New System.IO.FileInfo(pdfPath).Length
                Dim bytesPerPage As Double = If(pageCount > 0, CDbl(fileLen) / CDbl(pageCount), CDbl(fileLen))

                Dim textLen As Integer = If(extractedText IsNot Nothing, extractedText.Length, 0)

                ' Analyze text quality (document-level)
                Dim alphaNumCount As Integer = 0
                Dim whiteCount As Integer = 0
                Dim controlLikeCount As Integer = 0

                For Each ch As Char In extractedText
                    If System.Char.IsWhiteSpace(ch) Then
                        whiteCount += 1
                    End If
                    If System.Char.IsLetterOrDigit(ch) Then
                        alphaNumCount += 1
                    End If
                    If System.Char.IsControl(ch) AndAlso ch <> Microsoft.VisualBasic.ChrW(10) AndAlso ch <> Microsoft.VisualBasic.ChrW(13) AndAlso ch <> Microsoft.VisualBasic.ChrW(9) Then
                        controlLikeCount += 1
                    End If
                Next

                Dim alphaRatio As Double = If(textLen > 0, CDbl(alphaNumCount) / CDbl(textLen), 0.0)
                Dim whiteRatio As Double = If(textLen > 0, CDbl(whiteCount) / CDbl(textLen), 1.0)
                Dim controlRatio As Double = If(textLen > 0, CDbl(controlLikeCount) / CDbl(textLen), 0.0)

                Dim lettersPerPage As Double = If(pageCount > 0, CDbl(totalLetters) / CDbl(pageCount), 0.0)

                ' Threshold constants (document-level)
                Const MIN_TEXT_LEN_FOR_CONFIDENCE As Integer = 200
                Const MIN_LETTERS_PER_PAGE As Double = 15.0
                Const HIGH_BYTES_PER_PAGE As Double = 90_000
                Const LOW_ALPHA_RATIO As Double = 0.2
                Const HIGH_WHITE_RATIO As Double = 0.55
                Const HIGH_CONTROL_RATIO As Double = 0.02
                Const MANY_PAGES_FEW_LETTERS_PAGE_THRESHOLD As Integer = 5

                ' Page-level rules (strong signals)
                If pagesImageOnly.Count > 0 Then
                    shouldSuggestOcr = True
                    reasons.Add($"Found {pagesImageOnly.Count} image-only page(s) (0 text, 0 letters), e.g., page {pagesImageOnly(0)}.")
                ElseIf pagesLowLetters.Count > 0 Then
                    shouldSuggestOcr = True
                    reasons.Add($"Found {pagesLowLetters.Count} page(s) with very few letters (e.g., page {pagesLowLetters(0)}).")
                End If

                ' Document-level rules
                ' Rule A: Empty/near-empty text and large images per page
                If Not shouldSuggestOcr AndAlso textLen < MIN_TEXT_LEN_FOR_CONFIDENCE AndAlso bytesPerPage >= HIGH_BYTES_PER_PAGE Then
                    shouldSuggestOcr = True
                    reasons.Add($"Low extracted text ({textLen} chars) but large size per page (~{CInt(bytesPerPage)} bytes/page).")
                End If

                ' Rule B: very few letters per page on average
                If Not shouldSuggestOcr AndAlso lettersPerPage < MIN_LETTERS_PER_PAGE Then
                    shouldSuggestOcr = True
                    reasons.Add($"Very few letters detected by text layer on average (≈{lettersPerPage:N1} letters/page).")
                End If

                ' Rule C: Extracted text looks like junk (mostly whitespace/control or very low alpha)
                If Not shouldSuggestOcr AndAlso textLen > 0 AndAlso (alphaRatio < LOW_ALPHA_RATIO OrElse whiteRatio > HIGH_WHITE_RATIO OrElse controlRatio > HIGH_CONTROL_RATIO) Then
                    shouldSuggestOcr = True
                    reasons.Add($"Extracted text looks low-quality (alphaRatio={alphaRatio:P0}, whitespaceRatio={whiteRatio:P0}, controlRatio={controlRatio:P1}).")
                End If

                ' Rule D: Many pages but few letters overall
                If Not shouldSuggestOcr AndAlso pageCount >= MANY_PAGES_FEW_LETTERS_PAGE_THRESHOLD AndAlso lettersPerPage < MIN_LETTERS_PER_PAGE Then
                    shouldSuggestOcr = True
                    reasons.Add($"Many pages ({pageCount}) with very low letters/page (≈{lettersPerPage:N1}).")
                End If

                Debug.WriteLine($"PDF '{pdfPath}': pages={pageCount}, bytesPerPage≈{CInt(bytesPerPage)}, textLen={textLen}, lettersPerPage≈{lettersPerPage:N1}, alphaRatio={alphaRatio:P0}, whitespaceRatio={whiteRatio:P0}, controlRatio={controlRatio:P1}.")
                If shouldSuggestOcr Then
                    Debug.WriteLine("Heuristics suggest OCR. Reasons: " & String.Join(" | ", reasons.ToArray()))
                Else
                    Debug.WriteLine("Heuristics do not suggest OCR.")
                End If

                If shouldSuggestOcr Then
                    If AskUser Then
                        Dim formattedReasons As String = String.Join(Environment.NewLine, reasons.ConvertAll(Function(r) "- " & r))
                        Dim msg As String = $"The PDF appears to contain little or no extractable text:" & Environment.NewLine & Environment.NewLine &
                                            formattedReasons & Environment.NewLine & Environment.NewLine &
                                            "It's likely that the document consists mainly of scanned images." & Environment.NewLine & Environment.NewLine &
                                            "Would you like AI to perform OCR to extract text (if supported by your configured model)?"
                        Dim userChoice As Integer = ShowCustomYesNoBox(msg, "Yes, try OCR", "No, use what you have")
                        If userChoice <> 1 Then
                            Return extractedText
                        End If
                    End If

                    Dim ocrText As String = Await PerformOCR(pdfPath, context)
                    If Not String.IsNullOrWhiteSpace(ocrText) Then
                        Return ocrText
                    End If
                End If

                Return extractedText

            Catch ex As System.Exception
                Return If(ReturnErrorInsteadOfEmpty, $"Error reading PDF: {ex.Message}", "")
            End Try
        End Function


        Private Shared Function ExtractPageTextFromPdf(page As UglyToad.PdfPig.Content.Page) As String
            ' 1) Try PdfPig’s content-order extractor (good spacing/reading order on many PDFs)
            Try
                Dim t As String = UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor.ContentOrderTextExtractor.GetText(page)
                If Not String.IsNullOrWhiteSpace(t) AndAlso (t.Contains(" ") OrElse t.Contains(vbTab) OrElse t.Contains(vbCr) OrElse t.Contains(vbLf)) Then
                    Return t
                End If
            Catch
                ' Older PdfPig versions or certain pages may not support this path; ignore and fallback.
            End Try

            ' 2) Word-based reconstruction using Nearest-Neighbour extractor (higher recall on tricky PDFs)
            Try
                Dim words As System.Collections.Generic.IEnumerable(Of UglyToad.PdfPig.Content.Word) =
            page.GetWords(UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor.NearestNeighbourWordExtractor.Instance)

                If words IsNot Nothing AndAlso words.Count > 0 Then
                    ' Group words into lines by baseline with a tolerant threshold
                    Dim baselineTol As Double = Math.Max(0.5, page.Height * 0.002) ' ~0.2% of page height
                    Dim lines As New System.Collections.Generic.List(Of System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word))()

                    For Each w In words.OrderByDescending(Function(x) x.BoundingBox.Bottom).ThenBy(Function(x) x.BoundingBox.Left)
                        Dim placed As Boolean = False
                        For Each ln In lines
                            Dim ref = ln(0)
                            If Math.Abs(w.BoundingBox.Bottom - ref.BoundingBox.Bottom) <= baselineTol Then
                                ln.Add(w)
                                placed = True
                                Exit For
                            End If
                        Next
                        If Not placed Then
                            lines.Add(New System.Collections.Generic.List(Of UglyToad.PdfPig.Content.Word) From {w})
                        End If
                    Next

                    Dim sbLine As New System.Text.StringBuilder()
                    Dim first As Boolean = True
                    For Each ln In lines.OrderByDescending(Function(l) l.Average(Function(w) w.BoundingBox.Bottom))
                        If Not first Then sbLine.AppendLine()
                        first = False
                        Dim lineText = String.Join(" ", ln.OrderBy(Function(w) w.BoundingBox.Left).Select(Function(w) w.Text))
                        sbLine.Append(lineText)
                    Next

                    Dim s = sbLine.ToString()
                    If Not String.IsNullOrWhiteSpace(s) Then
                        Return s
                    End If
                End If
            Catch
                ' Ignore and fallback
            End Try

            ' 3) Letter-gap heuristic: insert spaces based on horizontal gaps; break lines on baseline changes
            Dim letters = page.Letters
            If letters Is Nothing OrElse letters.Count = 0 Then Return String.Empty

            Dim ordered = letters.OrderByDescending(Function(l) l.GlyphRectangle.Bottom).ThenBy(Function(l) l.GlyphRectangle.Left)
            Dim sb As New System.Text.StringBuilder()
            Dim prev As UglyToad.PdfPig.Content.Letter = Nothing

            For Each l In ordered
                If prev IsNot Nothing Then
                    Dim sameLine = Math.Abs(l.GlyphRectangle.Bottom - prev.GlyphRectangle.Bottom) <= Math.Max(0.5, prev.GlyphRectangle.Height * 0.6)
                    If Not sameLine Then
                        sb.AppendLine()
                    Else
                        Dim gap = l.GlyphRectangle.Left - prev.GlyphRectangle.Right
                        Dim spaceThreshold = Math.Max(prev.GlyphRectangle.Width * 0.6, 0.5) ' tune if needed
                        If gap > spaceThreshold Then sb.Append(" ")
                    End If
                End If
                sb.Append(l.Value)
                prev = l
            Next

            Return sb.ToString()
        End Function

        Private Shared Async Function PerformOCR(ByVal pdfPath As String, context As ISharedContext) As Task(Of String)

            If System.String.IsNullOrWhiteSpace(context.INI_APICall_Object) Then
                ShowCustomMessageBox($"Your model ({context.INI_Model}) is not configured to process binary objects - aborting OCR.")
                Return ""
            End If

            Dim UseSecondAPI As Boolean = False
            Dim TimeOut = context.INI_Timeout

            If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                If Not GetSpecialTaskModel(context, context.INI_AlternateModelPath, "OCR") Then
                    originalConfigLoaded = False
                    UseSecondAPI = False
                Else
                    UseSecondAPI = True
                    TimeOut = context.INI_Timeout_2
                End If
            End If

            Dim result As System.String = Await LLM(context, context.SP_InsertClipboard, "", "", "", TimeOut * 2, UseSecondAPI, False, "", pdfPath)

            ' Restore model if temporarily switched
            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(context, originalConfig)
                originalConfigLoaded = False
            End If

            Return result

        End Function


    End Class

End Namespace