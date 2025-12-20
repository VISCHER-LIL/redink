' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports HtmlAgilityPack
Imports Microsoft.Office.Interop.Word

Partial Public Class ThisAddIn

    Private Shared emojiSet As New HashSet(Of String)()
    Private Shared ReadOnly _emojiPairRegex As New System.Text.RegularExpressions.Regex(
    "[\uD83C-\uDBFF][\uDC00-\uDFFF]",
    System.Text.RegularExpressions.RegexOptions.Compiled Or
    System.Text.RegularExpressions.RegexOptions.CultureInvariant)


    ''' <summary>
    ''' Verknüpft zwei Formatierungs‑Delegates, ohne die ursprünglichen zu verlieren.
    ''' </summary>
    Private Shared Function CombineStyle(
    baseAction As Action(Of Microsoft.Office.Interop.Word.Range),
    additional As Action(Of Microsoft.Office.Interop.Word.Range)
) As Action(Of Microsoft.Office.Interop.Word.Range)

        If baseAction Is Nothing Then Return additional
        If additional Is Nothing Then Return baseAction

        Return Sub(rng As Microsoft.Office.Interop.Word.Range)
                   baseAction(rng)
                   additional(rng)
               End Sub
    End Function


    ''' <summary>
    ''' Rendert (ggf. rekursiv) reine Inline‑Knoten mit kumulativer Formatierung.
    ''' </summary>
    Private Shared Sub RenderInline(
    node As HtmlAgilityPack.HtmlNode,
    rng As Microsoft.Office.Interop.Word.Range,
    styleAction As Action(Of Microsoft.Office.Interop.Word.Range),
    inheritedHref As String
)

        ' Kommentare ignorieren
        If node.NodeType = HtmlAgilityPack.HtmlNodeType.Comment Then Return

        ' ------------------------------------------------- Leaf: #text -------------
        'If node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
        'Dim txt As String = HtmlEntity.DeEntitize(node.InnerText)
        'If Not String.IsNullOrWhiteSpace(txt) Then
        'InsertInline(rng, txt, styleAction, inheritedHref)
        'End If
        'Return
        'End If

        ' ─── Newline‑Handling: Nur echte, mit Inhalt versehene Zeilenumbrüche splitten
        If node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
            Dim rawText = HtmlEntity.DeEntitize(node.InnerText)
            Dim hasNewline = rawText.IndexOfAny({vbCr(0), vbLf(0)}) >= 0
            Dim stripped = rawText.Replace(vbCr, "").Replace(vbLf, "")

            ' 1) reine Whitespace‑-only‑Zeilenumbrüche → komplett überspringen
            If hasNewline AndAlso String.IsNullOrWhiteSpace(stripped) Then
                Return
            End If

            ' 2) gemischter Text mit echten Newlines → splitten & umbruchsweise einfügen
            If hasNewline Then
                Dim parts = rawText.Split(
            New String() {vbCrLf, vbCr, vbLf},
            StringSplitOptions.None)
                For i = 0 To parts.Length - 1
                    Dim segment = parts(i)
                    If Not String.IsNullOrWhiteSpace(segment) Then
                        InsertInline(rng, segment, styleAction, inheritedHref)
                    End If
                    If i < parts.Length - 1 Then
                        rng.InsertAfter(vbCr)
                        rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    End If
                Next
                Return
            End If

            ' 3) kein Newline → ganz normal einfügen
            If Not String.IsNullOrWhiteSpace(rawText) Then
                InsertInline(rng, rawText, styleAction, inheritedHref)
            End If
            Return
        End If


        ' ------------------------------------------------- Leaf: <br> --------------
        'If node.Name.Equals("br", StringComparison.OrdinalIgnoreCase) Then
        'rng.Font.Reset()
        'rng.Text = vbCr
        'rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        'Return
        'End If

        If node.Name.Equals("br", StringComparison.OrdinalIgnoreCase) Then
            ' Soft line break in Word (Shift+Enter) statt harter Absatz
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            rng.InsertAfter(ChrW(11))  ' manueller Zeilenumbruch
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            Return
        End If

        ' ------------------------------------------------- Leaf: <img> -------------
        If node.Name.Equals("img", StringComparison.OrdinalIgnoreCase) Then
            Dim src As String = node.GetAttributeValue("src", String.Empty)
            If Not String.IsNullOrWhiteSpace(src) Then
                ' Statt InlineShapes.AddPicture direkt aufzurufen,
                ' ruf den robusten Helper auf:
                InsertImageFromSrc(rng, src)
            End If
            Return
        End If

        ' ------------------------------------------------- Leaf/Semi‑Leaf: <a> -----
        Dim thisHref As String = inheritedHref
        If node.Name.Equals("a", StringComparison.OrdinalIgnoreCase) Then
            thisHref = node.GetAttributeValue("href", String.Empty)

            ' einfache Textlinks direkt ausgeben …
            If node.ChildNodes.All(Function(c) c.NodeType = HtmlAgilityPack.HtmlNodeType.Text) Then
                Dim txtLink As String = HtmlEntity.DeEntitize(node.InnerText)
                InsertInline(rng, txtLink, styleAction, thisHref)
                Return
            End If
            ' … ansonsten werden die Kinder rekursiv mit demselben href gerendert
        End If

        ' ------------------------------------------------- Style‑Weiche ------------
        Select Case node.Name.ToLowerInvariant()
            Case "strong", "b"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Bold = True)

            Case "em", "i"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Italic = True)

            Case "u"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Underline = Word.WdUnderline.wdUnderlineSingle)

            Case "del", "s"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.StrikeThrough = True)

            Case "code"
                styleAction = CombineStyle(styleAction,
                Sub(r)
                    r.Font.Name = "Courier New"
                    r.Font.Size = 10
                    r.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25
                End Sub)

            Case "sub"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Subscript = True)

            Case "sup"
                styleAction = CombineStyle(styleAction,
                                       Sub(r) r.Font.Superscript = True)

            Case "span"
                Dim cls = node.GetAttributeValue("class", String.Empty)
                If cls.Contains("emoji") Then
                    styleAction = CombineStyle(styleAction,
                    Sub(r)
                        r.Font.Name = "Segoe UI Emoji"
                        r.Font.Color = Word.WdColor.wdColorWhite
                        r.Shading.BackgroundPatternColor =
                            System.Drawing.ColorTranslator.ToOle(
                                System.Drawing.Color.FromArgb(0, 112, 192))
                    End Sub)
                End If
                ' sonst keine eigene Formatierung → einfach durchreichen
        End Select

        ' ------------------------------------------------- Rekursion ---------------
        For Each child In node.ChildNodes
            RenderInline(child, rng, styleAction, thisHref)
        Next
    End Sub




    ''' <summary>
    ''' Fügt ein Bild aus src in den Word-Range ein.
    ''' Unterstützt lokale Pfade und Web-URLs, und fängt alle Fehler intern ab.
    ''' </summary>
    Private Shared Sub InsertImageFromSrc(
    rng As Microsoft.Office.Interop.Word.Range,
    src As String
)
        If String.IsNullOrWhiteSpace(src) Then Return

        Dim fileName As String = src
        Dim tempFile As String = String.Empty
        Dim isUrl As Boolean = False

        Try
            ' URL-Erkennung
            Dim uri = New System.Uri(src, UriKind.RelativeOrAbsolute)
            If uri.IsAbsoluteUri AndAlso
           (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) _
            OrElse uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) Then

                isUrl = True
                tempFile = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                System.IO.Path.GetFileName(uri.LocalPath)
            )
                Using client As New System.Net.WebClient()
                    client.DownloadFile(uri, tempFile)
                End Using
                fileName = tempFile
            End If

            ' Existenz prüfen
            If Not System.IO.File.Exists(fileName) Then
                Throw New System.Exception(
                $"Bilddatei nicht gefunden bzw. Download fehlgeschlagen: '{fileName}'"
            )
            End If

            ' Einfügen
            Dim pic As Microsoft.Office.Interop.Word.InlineShape =
            rng.InlineShapes.AddPicture(
                FileName:=fileName,
                LinkToFile:=False,
                SaveWithDocument:=True
            )
            rng.SetRange(pic.Range.End, pic.Range.End)

        Catch ex As System.Exception
            ' Fehler intern loggen, nicht weiterwerfen
            Debug.WriteLine(
            $"[InsertImageFromSrc] {ex.GetType().FullName}: {ex.Message}"
        )
            rng.InsertAfter("[Image missing]")
        Finally
            ' Temp-Datei entfernen
            If isUrl AndAlso Not String.IsNullOrWhiteSpace(tempFile) Then
                Try
                    System.IO.File.Delete(tempFile)
                Catch ioEx As System.Exception
                    Debug.WriteLine(
                    $"[InsertImageFromSrc] Temp-Datei konnte nicht gelöscht werden: {ioEx.Message}"
                )
                End Try
            End If
        End Try
    End Sub

    Private Shared Sub RemoveTrailingParagraph(htmlDoc As HtmlAgilityPack.HtmlDocument)

        Dim candidates = New String() {"p", "br"}

        For Each TagName In candidates
            Dim lastNode = htmlDoc.DocumentNode.SelectSingleNode("(//" & TagName & ")[last()]")
            If lastNode Is Nothing Then Continue For

            ' liegt wirklich ganz am Schluss?
            Dim cur = lastNode.NextSibling
            Dim canDelete As Boolean = True
            While cur IsNot Nothing
                Select Case cur.NodeType
                    Case HtmlAgilityPack.HtmlNodeType.Comment
                    ' ignorieren
                    Case HtmlAgilityPack.HtmlNodeType.Text
                        If Not String.IsNullOrWhiteSpace(cur.InnerText) Then
                            canDelete = False : Exit While
                        End If
                    Case Else
                        canDelete = False : Exit While
                End Select
                cur = cur.NextSibling
            End While

            If canDelete Then
                If TagName = "p" Then
                    ' Kinder retten, <p> selbst raus
                    For Each child In lastNode.ChildNodes.ToList()
                        lastNode.ParentNode.InsertBefore(child, lastNode)
                    Next
                End If
                lastNode.Remove()
                Exit For            ' nur EIN Abschlusstag anfassen
            End If
        Next
    End Sub




    Private Shared ulLevels As List(Of Integer)
    Private Shared ulStartPos As Integer

    Private Shared Sub ParseHtmlNode(
    node As HtmlAgilityPack.HtmlNode,
    range As Microsoft.Office.Interop.Word.Range,
    Optional currentLevel As Integer = 0)

        ' -------------------------------- Text‑Shortcut ---------------------------
        If Not node.HasChildNodes AndAlso node.NodeType = HtmlAgilityPack.HtmlNodeType.Text Then
            RenderInline(node, range, Nothing, String.Empty)
            Return
        End If


        If node.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then
            ' (1) Inline‐Inhalt rendern
            RenderInline(node, range, Nothing, String.Empty)

            ' — nächstes echtes Node finden (Whitespace/Comments überspringen) —
            Dim nxt As HtmlAgilityPack.HtmlNode = node.NextSibling
            While nxt IsNot Nothing _
            AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                    OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                            AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                nxt = nxt.NextSibling
            End While

            ' — nur wenn wirklich noch etwas folgt, Absatz(e) einfügen —
            If nxt IsNot Nothing Then
                ' 1× Absat­zumbruch
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' noch ein Leerabsatz, falls nächstes Geschwister ein <p> ist
                If nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                End If
            End If

            Return
        End If


        If node.Name.Equals("li", StringComparison.OrdinalIgnoreCase) Then

            Dim isFootnoteEntry As Boolean =
                    node.GetAttributeValue("id", String.Empty) _
                        .StartsWith("fn:", StringComparison.OrdinalIgnoreCase)

            If isFootnoteEntry Then
                ' --- (A) Bookmark‑Name und Fußnotenzahl ermitteln ---
                Dim rawId As String = node.GetAttributeValue("id", String.Empty)   ' z.B. "fn:1"
                Dim fnNum As String = rawId.Substring(rawId.IndexOf(":"c) + 1)     ' "1"
                Dim bookmarkName As String = "fn" & fnNum                          ' "fn1" (gültiger Bookmark-Name)

                ' --- (B) Superscript-Zahl einfügen und Bookmark um sie herum anlegen ---
                Dim bmStart As Integer = range.End
                InsertInline(
                                range,
                                fnNum,
                                Sub(r) r.Font.Superscript = True,
                                String.Empty
                            )
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                'Dim bmRange As Word.Range = range.Document.Range(bmStart, range.End)
                Dim bmRange As Word.Range = range.Duplicate
                bmRange.Start = bmStart
                bmRange.End = range.End
                range.Bookmarks.Add(Name:=bookmarkName, Range:=bmRange)

                Debug.WriteLine($"[ParseHtmlNode] Footnote Bookmark '{bookmarkName}' at Range=({bmStart},{range.End})")

                ' Leerzeichen nach der Zahl
                range.InsertAfter(" ")
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' --- (C) Den gesamten Fußnoten-Text (und Rücksprung-Arrow) rendern ---
                ' Falls das <li> nur ein <p> enthält, entpacken wir es
                If node.ChildNodes.Count = 1 _
                     AndAlso node.FirstChild.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then

                    Dim pNode As HtmlAgilityPack.HtmlNode = node.FirstChild
                    For Each subNode As HtmlAgilityPack.HtmlNode In pNode.ChildNodes
                        ParseHtmlNode(subNode, range, currentLevel)
                    Next

                Else

                    For Each subNode As HtmlAgilityPack.HtmlNode In node.ChildNodes

                        Select Case subNode.Name.ToLowerInvariant()

                            Case "p"
                                ' (1) Inline-Inhalt rendern (inkl. <br> als manueller Break)
                                RenderInline(subNode, range, Nothing, String.Empty)

                                ' (2) echten Absatz abschließen
                                range.InsertParagraphAfter()
                                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                ' (3) prüfen, ob direkt ein <p> oder <blockquote> folgt (nach Whitespace/Comments)
                                Dim nxt As HtmlAgilityPack.HtmlNode = subNode.NextSibling
                                While nxt IsNot Nothing _
                                          AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                                   OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                                           AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                                    nxt = nxt.NextSibling
                                End While

                                ' (4) wenn ja, noch eine leere Zeile einfügen
                                If nxt IsNot Nothing _
                                       AndAlso (nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) _
                                                OrElse nxt.Name.Equals("blockquote", StringComparison.OrdinalIgnoreCase)) Then

                                    range.InsertParagraphAfter()
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                End If

                                Exit Select


                            Case "blockquote"
                                ' (1) Zitat‑Absätze einzeln rendern
                                Dim quoteParas = subNode.SelectNodes("./p")
                                If quoteParas IsNot Nothing Then
                                    For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas

                                        Dim paraStart As Integer = range.Start
                                        RenderInline(pNode, range, Nothing, String.Empty)

                                        range.InsertParagraphAfter()
                                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                        ' Den soeben eingefügten Absatz einrücken
                                        Dim indentRg As Microsoft.Office.Interop.Word.Range = range.Duplicate
                                        indentRg.Start = paraStart
                                        indentRg.End = range.End
                                        'range.Document.Range(paraStart, range.End)

                                        indentRg.ListFormat.RemoveNumbers()
                                        indentRg.ParagraphFormat.LeftIndent +=
                                            indentRg.Application.CentimetersToPoints(0.75)


                                    Next
                                Else
                                    ' Fallback: ohne <p>-Wrapper rekursiv parsen
                                    For Each innerNode As HtmlAgilityPack.HtmlNode In subNode.ChildNodes
                                        ParseHtmlNode(innerNode, range, currentLevel)
                                    Next
                                End If

                                Exit Select


                        ' Inline‑Elemente direkt rendern, damit RenderInline den Fett‑Stil anwendet:
                            Case "#text", "strong", "b", "em", "i", "u",
                             "del", "s", "sub", "sup", "code", "span", "img", "br", "a"

                                RenderInline(subNode, range, Nothing, String.Empty)

                        ' verschachtelte Listen wie gehabt überspringen
                            Case "ul", "ol"
                                ' nichts tun, wird unten separat behandelt

                                ' alle anderen Block‑Elemente rekursiv parsen
                            Case Else
                                ParseHtmlNode(subNode, range, currentLevel)

                        End Select

                    Next

                End If

                ' --- (D) Absatz nach jeder Fußnote und Rückkehr ---
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Return
            End If

            If currentLevel > 1 Then
                ' --- (0) CR VOR dem LI, je nach Listen‑Typ ---
                Dim parentName = node.ParentNode.Name.ToLowerInvariant()
                If parentName = "ol" Or parentName = "ul" Then
                    ' OL: nur vor Unterpunkten außer dem ersten
                    Dim sibs = node.ParentNode.SelectNodes("li")
                    If sibs IsNot Nothing Then
                        Dim idx As Integer = 0
                        For i As Integer = 0 To sibs.Count - 1
                            If sibs(i) Is node Then
                                idx = i
                                Exit For
                            End If
                        Next
                        If idx > 0 Then
                            range.InsertAfter(vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If
                    End If
                End If
            End If

            ' --- (1) Level speichern ---
            If ulLevels IsNot Nothing Then
                ulLevels.Add(currentLevel)
            End If

            ' --- (2) P‑Wrapper entfernen, wenn er das einzige direkte Kind ist ---
            If node.ChildNodes.Count = 1 _
       AndAlso node.FirstChild.Name.Equals("p", StringComparison.OrdinalIgnoreCase) Then

                Dim pNode As HtmlAgilityPack.HtmlNode = node.FirstChild
                For Each subNode As HtmlAgilityPack.HtmlNode In pNode.ChildNodes
                    ParseHtmlNode(subNode, range, currentLevel)
                Next

            Else

                For Each subNode As HtmlAgilityPack.HtmlNode In node.ChildNodes

                    Select Case subNode.Name.ToLowerInvariant()

                        Case "blockquote"
                            ' (1) Zitat‑Absätze einzeln rendern
                            Dim quoteParas = subNode.SelectNodes("./p")
                            If quoteParas IsNot Nothing Then
                                For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas
                                    Dim paraStart As Integer = range.Start
                                    RenderInline(pNode, range, Nothing, String.Empty)

                                    range.InsertParagraphAfter()
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                                    ' Den soeben eingefügten Absatz einrücken
                                    Dim indentRg As Microsoft.Office.Interop.Word.Range = range.Duplicate
                                    indentRg.Start = paraStart
                                    indentRg.End = range.End
                                    indentRg.ListFormat.RemoveNumbers()
                                    indentRg.ParagraphFormat.LeftIndent +=
                                        indentRg.Application.CentimetersToPoints(0.75)


                                Next
                            Else
                                ' Fallback: ohne <p>-Wrapper rekursiv parsen
                                For Each innerNode As HtmlAgilityPack.HtmlNode In subNode.ChildNodes
                                    ParseHtmlNode(innerNode, range, currentLevel)
                                Next
                            End If

                            Exit Select

                        ' Inline‑Elemente direkt rendern, damit RenderInline den Fett‑Stil anwendet:
                        Case "#text", "strong", "b", "em", "i", "u",
                             "del", "s", "sub", "sup", "code", "span", "img", "br", "a"

                            RenderInline(subNode, range, Nothing, String.Empty)

                        ' verschachtelte Listen wie gehabt überspringen
                        Case "ul", "ol"
                            ' nichts tun, wird unten separat behandelt

                            ' alle anderen Block‑Elemente rekursiv parsen
                        Case Else
                            ParseHtmlNode(subNode, range, currentLevel)

                    End Select

                Next

            End If

            ' --- (3) Verschachtelte Listen am Ende behandeln ---
            Dim nestedUl As HtmlAgilityPack.HtmlNode = node.SelectSingleNode("ul")
            Dim nestedOl As HtmlAgilityPack.HtmlNode = node.SelectSingleNode("ol")
            If (nestedUl IsNot Nothing OrElse nestedOl IsNot Nothing) Then
                range.InsertAfter(vbCr)
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If
            If nestedUl IsNot Nothing Then
                ParseHtmlNode(nestedUl, range, currentLevel + 1)
            ElseIf nestedOl IsNot Nothing Then
                ParseHtmlNode(nestedOl, range, currentLevel + 1)
            End If

            If isFootnoteEntry Then
                range.InsertParagraphAfter()
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            End If

            Return
        End If


        Debug.WriteLine($"[ParseHtmlNode] Enter node=<{node.Name}> Range=({range.Start},{range.End})")

        ' -------------------------------- Haupt, und Kindknoten‑Schleife ---------------------
        For Each childNode As HtmlAgilityPack.HtmlNode In node.ChildNodes


            Debug.WriteLine($"  └─ Child: <{childNode.Name}> Type={childNode.NodeType}")

            Dim nestedLinkNode As HtmlAgilityPack.HtmlNode = Nothing
            If Not childNode.Name.Equals("a", StringComparison.OrdinalIgnoreCase) Then
                nestedLinkNode = childNode.SelectSingleNode(".//a")
            End If
            Dim nestedHref As String = If(nestedLinkNode IsNot Nothing,
                                      nestedLinkNode.GetAttributeValue("href", String.Empty),
                                      String.Empty)

            Select Case childNode.Name.ToLowerInvariant()


                Case "blockquote"
                    ' (1) Absatz vor dem Zitat
                    'range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    'range.InsertParagraphAfter()
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                    ' (2) Nur direkte <p>-Kinder verarbeiten
                    Dim quoteParas As HtmlAgilityPack.HtmlNodeCollection =
                        childNode.SelectNodes("./p")

                    If quoteParas IsNot Nothing Then
                        For Each pNode As HtmlAgilityPack.HtmlNode In quoteParas
                            ' Markiere den Anfang des neuen Absatzes
                            Dim paraStart As Integer = range.Start

                            ' (3) Inline-Inhalt des Zitats rendern
                            RenderInline(pNode, range, Nothing, String.Empty)

                            ' (4) Absatz nach jedem Zitat-Absatz
                            range.InsertParagraphAfter()
                            range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                            ' (5) Den soeben eingefügten Absatz einrücken
                            Dim indentRg As Microsoft.Office.Interop.Word.Range = range.Duplicate
                            indentRg.Start = paraStart
                            indentRg.End = range.End
                            indentRg.ListFormat.RemoveNumbers()
                            indentRg.ParagraphFormat.LeftIndent +=
                                indentRg.Application.CentimetersToPoints(0.75)

                        Next
                    Else
                        ' Fallback: kein <p> im Blockquote → normal parsen
                        For Each subNode As HtmlAgilityPack.HtmlNode In childNode.ChildNodes
                            ParseHtmlNode(subNode, range, currentLevel)
                        Next
                    End If

                    Exit Select

                    ' (6) Absatz nach dem gesamten Blockquote
                    range.InsertParagraphAfter()
                    range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                    ' (7) zusätzlichen Leerabsatz nur, wenn nach </blockquote> direkt ein <p> folgt
                    Dim nxtBQ As HtmlAgilityPack.HtmlNode = childNode.NextSibling
                    While nxtBQ IsNot Nothing _
                              AndAlso (nxtBQ.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                       OrElse (nxtBQ.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                               AndAlso String.IsNullOrWhiteSpace(nxtBQ.InnerText)))
                        nxtBQ = nxtBQ.NextSibling
                    End While

                    If nxtBQ IsNot Nothing _
                           AndAlso nxtBQ.Name.Equals("p", System.StringComparison.OrdinalIgnoreCase) Then

                        range.InsertParagraphAfter()
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                    End If

                    Exit Select

                Case "p"
                    ' Inline-Inhalt rendern (inkl. <br> als manueller Umbruch)
                    RenderInline(childNode, range, Nothing, String.Empty)

                    ' nächstes echtes Geschwister-Node finden
                    Dim nxt As HtmlAgilityPack.HtmlNode = childNode.NextSibling
                    While nxt IsNot Nothing _
                          AndAlso (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Comment _
                                   OrElse (nxt.NodeType = HtmlAgilityPack.HtmlNodeType.Text _
                                           AndAlso String.IsNullOrWhiteSpace(nxt.InnerText)))
                        nxt = nxt.NextSibling
                    End While

                    ' nur wenn wirklich etwas folgt, 1× Absatz
                    If nxt IsNot Nothing Then
                        range.InsertParagraphAfter()
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        ' zusätzlich eine Leerzeile, wenn das Geschwister ein weiterer <p> ist
                        If nxt.Name.Equals("p", StringComparison.OrdinalIgnoreCase) OrElse nxt.Name.Equals("blockquote", StringComparison.OrdinalIgnoreCase) Then
                            range.InsertParagraphAfter()
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If
                    End If

                    Exit Select



                Case "#text", "strong", "b", "em", "i", "u", "del", "s",
                    "sub", "sup", "code", "span", "img", "br"
                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "div"
                    Dim cls As String = childNode.GetAttributeValue("class", String.Empty)
                    If cls.Equals("footnotes", StringComparison.OrdinalIgnoreCase) Then
                        ' Statt zu überspringen: das OL innerhalb der Fußnoten parsen
                        Dim footOl As HtmlAgilityPack.HtmlNode = childNode.SelectSingleNode("ol")
                        If footOl IsNot Nothing Then
                            ' currentLevel evtl. gleich lassen oder auf 0 setzen,
                            ' je nachdem, wie du die Nummerierung haben willst
                            ParseHtmlNode(footOl, range, currentLevel)
                        End If
                        Exit Select
                    End If

                Case "br"
                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "h1", "h2", "h3", "h4", "h5", "h6"
                    ' 1) Welcher Built-In Heading-Style?

                    Dim style As WdBuiltinStyle = WdBuiltinStyle.wdStyleNormal ' Default for 'p'
                    Select Case childNode.Name.ToLower()
                        Case "h1" : style = WdBuiltinStyle.wdStyleHeading1
                        Case "h2" : style = WdBuiltinStyle.wdStyleHeading2
                        Case "h3" : style = WdBuiltinStyle.wdStyleHeading3
                        Case "h4" : style = WdBuiltinStyle.wdStyleHeading4
                        Case "h5" : style = WdBuiltinStyle.wdStyleHeading5
                        Case "h6" : style = WdBuiltinStyle.wdStyleHeading6
                    End Select

                    Dim txt As String = HtmlEntity.DeEntitize(childNode.InnerText)
                    Dim href As String = nestedHref

                    ' 2) Neuen Absatz einfügen und Range dorthin setzen
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ' 3) Text einfügen
                    Dim paraStart As Integer = range.Start
                    range.InsertAfter(txt)
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    ' 4) Absatz-Range ermitteln
                    Dim paraRg As Word.Range = range.Duplicate
                    paraRg.Start = paraStart
                    paraRg.End = range.End
                    ' 5) Absatz-Stil anwenden
                    paraRg.Style = style

                    ' 6) Hyperlink (falls nötig)
                    If href <> String.Empty Then
                        Dim hl As Word.Hyperlink =
                                            range.Hyperlinks.Add(
                                                Anchor:=paraRg,
                                                Address:=CStr(href)
                                            )
                        hl.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        range.SetRange(hl.Range.End, hl.Range.End)
                    End If

                    ' 7) Absatz-Umbruch ans Ende
                    range.InsertParagraphAfter()
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                Case "a"
                    Dim cls As String = childNode.GetAttributeValue("class", String.Empty)
                    Dim href As String = childNode.GetAttributeValue("href", String.Empty)
                    Dim id As String = childNode.GetAttributeValue("id", String.Empty)

                    ' 1) Inline-Fußnoten-Referenz im Text?
                    If id.StartsWith("fnref:", StringComparison.OrdinalIgnoreCase) _
                       AndAlso href.StartsWith("#fn:", StringComparison.OrdinalIgnoreCase) Then

                        Debug.WriteLine("Setting Bookmark in Case 'a'")

                        ' Fußnotenzahl extrahieren
                        Dim fnNum As String = id.Substring(id.IndexOf(":"c) + 1)  ' z.B. "1"
                        Dim bookmarkName As String = "fn" & fnNum                         ' "fn1"

                        ' Anzeige-Text (die <sup>1</sup>)
                        Dim displayText As String = HtmlEntity.DeEntitize(childNode.InnerText)

                        'Hyperlink auf unser Bookmark anlegen
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        range.Hyperlinks.Add(
                                                Anchor:=range,
                                                Address:="",                ' keine externe URL
                                                SubAddress:=CStr(bookmarkName),   ' internes Ziel
                                                TextToDisplay:=CStr(displayText)
                                                )
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                        Exit Select    ' <-- hier beenden wir NUR den Case, nicht die ganze Sub
                    End If

                    ' 2) Rücksprung-Pfeil in der Fußnotenliste
                    If cls.Equals("footnote-back-ref", StringComparison.OrdinalIgnoreCase) Then
                        RenderInline(childNode, range, Nothing, String.Empty)
                        Exit Select
                    End If

                    ' 3) alle anderen Links normal
                    RenderInline(childNode, range, Nothing, String.Empty)
                    Exit Select




                Case "ul"
                    ' a) Task‑List (Checkboxen) wie gehabt …
                    If childNode.GetAttributeValue("class", "").Contains("contains-task-list") Then
                        For Each li As HtmlAgilityPack.HtmlNode In childNode.SelectNodes("li")
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Dim chk = li.SelectSingleNode(".//input[@type='checkbox']")
                            Dim symbol = If(chk IsNot Nothing _
                           AndAlso chk.GetAttributeValue("checked", False), "☑", "☐")
                            range.InsertAfter(symbol & " " &
                              HtmlEntity.DeEntitize(li.InnerText.Trim()) &
                              vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Next
                        Exit Select
                    End If

                    Dim enteringTopUL = (currentLevel = 0)
                    If enteringTopUL AndAlso ulLevels Is Nothing Then
                        ulLevels = New List(Of Integer)()
                        ulStartPos = range.Start
                        Debug.WriteLine("[ul] Entering top UL – reset ulLevels")
                    End If

                    ' c) Originales Einfügen der LI‑Nodes (recursive)
                    Dim liNodes = childNode.SelectNodes("li")
                    If liNodes IsNot Nothing Then
                        Dim listStart = range.Start
                        For Each liNode As HtmlAgilityPack.HtmlNode In liNodes
                            ParseHtmlNode(liNode, range, currentLevel + 1)
                            range.InsertAfter(vbCr)
                            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Next
                        Dim ulRange As Word.Range = range.Duplicate
                        ulRange.Start = listStart   ' listStart must be in the same story as `range`
                        ulRange.End = range.End
                        ulRange.ListFormat.ApplyBulletDefault()
                        ulRange.ListFormat.ListIndent()
                        With ulRange.ParagraphFormat
                            .LeftIndent = .Application.CentimetersToPoints(0.75)
                            .FirstLineIndent = - .Application.CentimetersToPoints(0.75)
                        End With
                        range.SetRange(ulRange.End, ulRange.End)

                        ' d) Am Ende des ersten UL: Array anwenden
                        If enteringTopUL Then
                            Debug.WriteLine($"[ul] At end of top UL – ulLevels.Count = {ulLevels.Count}")
                            Debug.WriteLine($"[ul] Levels array: {String.Join(",", ulLevels)}")

                            Dim paras = ulRange.Paragraphs
                            Dim maxItems = System.Math.Min(paras.Count, ulLevels.Count)
                            For i = 1 To maxItems
                                Dim p = paras(i)
                                Dim lvl = ulLevels(i - 1)
                                Debug.WriteLine($"[ul] Paragraph {i} initial level={lvl}")

                                ' jede weitere Ebene einmal ListIndent()
                                For stepIndent = 1 To (lvl - 1)
                                    p.Range.ListFormat.ListIndent()
                                Next
                            Next
                            ulLevels = Nothing
                        End If
                    End If
                    If enteringTopUL Then
                        ulLevels = Nothing
                    End If

                    Exit Select


                Case "ol"

                    ' a) Alle <li>-Knoten holen
                    Dim liNodes As HtmlAgilityPack.HtmlNodeCollection = childNode.SelectNodes("li")
                    If liNodes Is Nothing OrElse liNodes.Count = 0 Then
                        Debug.WriteLine("[ol] No <li> nodes found in <ol> – skipping")
                        Exit Select
                    End If

                    ' b) Start-Attribut auslesen (Startnummer)
                    Dim startAttr As Integer = 1
                    Dim startStr As String = childNode.GetAttributeValue("start", String.Empty)
                    Dim tmpInt As Integer
                    If Integer.TryParse(startStr, tmpInt) Then startAttr = tmpInt

                    ' c) Top-Level-OL: ulLevels initialisieren (shared mit UL)
                    Dim enteringTopUL As Boolean = (currentLevel = 0)
                    If enteringTopUL AndAlso ulLevels Is Nothing Then
                        ulLevels = New List(Of Integer)()
                        ulStartPos = range.Start
                        Debug.WriteLine("[ol] Entering top OL – reset ulLevels")
                    End If

                    ' d) Jedes LI rekursiv rendern (die LI-Logik übernimmt CR und Level-Push)
                    Dim listStart As Integer = range.Start
                    For Each liNode As HtmlAgilityPack.HtmlNode In liNodes
                        ParseHtmlNode(liNode, range, currentLevel + 1)
                        range.InsertAfter(vbCr)
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Next

                    ' e) Den kompletten Bereich der OL-Liste erfassen
                    Dim olRange As Word.Range = range.Duplicate
                    olRange.Start = listStart
                    olRange.End = range.End

                    ' f) Prüfen, ob olRange überhaupt Absätze enthält
                    If olRange.Paragraphs.Count = 0 Then
                        Debug.WriteLine("[ol] olRange enthält keine Absätze – überspringe Nummerierung")
                        Exit Select
                    End If

                    ' j) Formatierung nur beim obersten OL (enteringTopUL) anwenden
                    If enteringTopUL Then
                        Dim paras As Word.Paragraphs = olRange.Paragraphs

                        ' Remove any previous numbering
                        olRange.ListFormat.RemoveNumbers()

                        ' Use a multi-level list template from ListGalleries
                        Dim multiLevelTemplate As Word.ListTemplate = olRange.Application.ListGalleries(Word.WdListGalleryType.wdOutlineNumberGallery).ListTemplates(1)

                        ' Set custom start value for level 1 if needed
                        If startAttr <> 1 Then
                            multiLevelTemplate.ListLevels(1).StartAt = startAttr
                        End If

                        ' Apply the multi-level template to the range
                        olRange.ListFormat.ApplyListTemplateWithLevel(
                                ListTemplate:=multiLevelTemplate,
                                ContinuePreviousList:=False,
                                ApplyTo:=Word.WdListApplyTo.wdListApplyToSelection,
                                DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior
                            )

                        ' Set the correct level for each paragraph
                        For i As Integer = 1 To paras.Count
                            Dim p As Word.Paragraph = paras(i)
                            Dim lvl As Integer = ulLevels(i - 1)
                            If lvl >= 1 AndAlso lvl <= multiLevelTemplate.ListLevels.Count Then
                                p.Range.ListFormat.ListLevelNumber = lvl
                            End If
                        Next
                        ulLevels = Nothing
                    End If

                    ' j) Cursor hinter die Liste setzen
                    range.SetRange(olRange.End, olRange.End)

                    Exit Select

                Case "dl"
                    ' Definition list
                    For Each dt As HtmlNode In childNode.SelectNodes("dt")
                        ' Term
                        Dim term As Microsoft.Office.Interop.Word.Range = range.Duplicate
                        term.Text = HtmlEntity.DeEntitize(dt.InnerText) & vbTab
                        term.Font.Bold = True
                        term.Collapse(False)
                        range.SetRange(term.End, term.End)
                        ' Definition
                        Dim dd As HtmlNode = dt.NextSibling
                        If dd IsNot Nothing AndAlso dd.Name.ToLower() = "dd" Then
                            Dim defn As Microsoft.Office.Interop.Word.Range = range.Duplicate
                            defn.Text = HtmlEntity.DeEntitize(dd.InnerText) & vbCr
                            defn.ParagraphFormat.LeftIndent += 18
                            defn.Collapse(False)
                            range.SetRange(defn.End, defn.End)
                        End If
                    Next


                Case "hr"

                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                    Dim hrPara As Word.Paragraph = range.Paragraphs.Add(range)
                    hrPara.Range.Text = ""  ' leer lassen, wir brauchen nur den Rahmen

                    With hrPara.Range.ParagraphFormat.Borders(Word.WdBorderType.wdBorderBottom)
                        .LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        .LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        .Color = Word.WdColor.wdColorAutomatic
                    End With

                    Dim afterHr As Word.Range = hrPara.Range.Duplicate
                    afterHr.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    range.SetRange(afterHr.Start, afterHr.Start)
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)



                Case "input"
                    ' Checkbox (ContentControl)
                    If childNode.GetAttributeValue("type", String.Empty).ToLower() = "checkbox" Then
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Dim cc As Word.ContentControl =
                                        range.ContentControls.Add(
                                            Word.WdContentControlType.wdContentControlCheckBox,
                                            range
                                        )
                        cc.Checked = (childNode.GetAttributeValue("checked", String.Empty).ToLower() = "checked")

                        range.SetRange(cc.Range.End, cc.Range.End)
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    End If

                Case "img"

                    RenderInline(childNode, range, Nothing, String.Empty)

                Case "pre"
                    ' Code block
                    Dim codeBlock As Microsoft.Office.Interop.Word.Range = range.Duplicate
                    codeBlock.Text = HtmlEntity.DeEntitize(childNode.InnerText) & vbCr
                    codeBlock.Font.Name = "Courier New"
                    codeBlock.Font.Size = 10
                    codeBlock.ParagraphFormat.LeftIndent += 14.18
                    codeBlock.Collapse(False)
                    range.SetRange(codeBlock.End, codeBlock.End)




                Case "table"
                    '---------- 1) Top-Level-Rows holen ----------------------------
                    Dim topRows As New List(Of HtmlNode)

                    'direkte <tr> plus <thead>/<tbody>-Kinder, aber KEINE rekursiven
                    For Each tr As HtmlNode In childNode.SelectNodes("./tr|./thead/tr|./tbody/tr")
                        topRows.Add(tr)
                    Next
                    If topRows.Count = 0 Then Exit Select
                    '----------------------------------------------------------------

                    '---------- 2) Beste Spaltenzahl ermitteln ----------------------
                    Dim colCount As Integer = 0
                    For Each tr In topRows
                        Dim cells = tr.SelectNodes("th|td")
                        If cells IsNot Nothing AndAlso cells.Count > colCount Then
                            colCount = cells.Count
                        End If
                    Next
                    If colCount = 0 Then Exit Select
                    '----------------------------------------------------------------

                    '---------- 3) Tabelle an Cursor anlegen ------------------------
                    Dim tbl As Microsoft.Office.Interop.Word.Table =
                        range.Tables.Add(range, topRows.Count, colCount)

                    '---------- 4) Zellen befüllen ----------------------------------
                    Dim rIdx As Integer = 1
                    For Each tr In topRows
                        Dim cells = tr.SelectNodes("th|td")
                        Dim cIdx As Integer = 1

                        If cells IsNot Nothing Then
                            For Each cell In cells
                                Dim cellRg As Word.Range = tbl.Cell(rIdx, cIdx).Range
                                'unsichtbares Zellenendzeichen abschneiden
                                cellRg.SetRange(cellRg.Start, cellRg.End - 1)

                                ParseHtmlNode(cell, cellRg, currentLevel)          '← rekursiv, kein Datenverlust

                                'Headerzelle fett
                                If cell.Name.Equals("th", StringComparison.OrdinalIgnoreCase) Then
                                    cellRg.Font.Bold = True
                                End If

                                '---------- 4a) Colspan behandeln ------------------
                                Dim cSpan As Integer = cell.GetAttributeValue("colspan", 1)
                                If cSpan > 1 AndAlso cIdx + cSpan - 1 <= colCount Then
                                    Dim tgtCell = tbl.Cell(rIdx, cIdx + cSpan - 1)
                                    tbl.Cell(rIdx, cIdx).Merge(tgtCell)
                                    cIdx += cSpan                    'gleich weiter hinter dem Merge
                                Else
                                    cIdx += 1
                                End If
                                '----------------------------------------------------
                            Next
                        End If
                        rIdx += 1
                    Next

                    '---------- 5) Cursor hinter Tabelle setzen --------------------
                    range.SetRange(tbl.Range.End, tbl.Range.End)


                Case Else

                    ParseHtmlNode(childNode, range, currentLevel)

            End Select

        Next
    End Sub





    Private Shared Sub InsertInline(
        ByRef mainRg As Range,
        txt As String,
        baseStyle As Action(Of Range),
        Optional href As String = "")

        ' 1) Kein Emoji‑Fall → direkter Einfügen‑Pfad
        If emojiSet Is Nothing OrElse emojiSet.Count = 0 OrElse Not _emojiPairRegex.IsMatch(txt) Then
            TrueInsertInline(mainRg, txt, baseStyle, href)
            Return
        End If

        ' 2) Sonst: nur an den Emoji-Punkten splitten und stylen
        Dim lastPos As Integer = 0
        For Each m As Match In _emojiPairRegex.Matches(txt)
            ' (a) Alles vor dem Emoji
            If m.Index > lastPos Then
                Dim segment As String = txt.Substring(lastPos, m.Index - lastPos)
                TrueInsertInline(mainRg, segment, baseStyle, href)
            End If

            ' (b) Das Emoji selbst (nur, wenn es im Set ist)
            Dim emoji As String = m.Value
            If emojiSet.Contains(emoji) Then
                Dim emojiStyle = CombineStyle(baseStyle,
                    Sub(r As Range) r.Font.Name = "Segoe UI Emoji")
                TrueInsertInline(mainRg, emoji, emojiStyle, href)
            Else
                ' falls doch nicht im Set – als normaler Text
                TrueInsertInline(mainRg, emoji, baseStyle, href)
            End If

            lastPos = m.Index + m.Length
        Next

        ' (c) Rest nach dem letzten Emoji
        If lastPos < txt.Length Then
            Dim tail As String = txt.Substring(lastPos)
            TrueInsertInline(mainRg, tail, baseStyle, href)
        End If
    End Sub




    Private Shared Sub TrueInsertInline(
    ByRef mainRg As Word.Range,
    txt As String,
    styleAction As Action(Of Word.Range),
    Optional href As String = "")

        mainRg.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

        Dim wrk As Word.Range = mainRg.Duplicate
        wrk.Text = txt

        ' **Hier passiert der Reset** – denk daran, vor und nach dem Reset zu loggen
        wrk.Font.Reset()

        If href <> "" Then
            Dim hl = mainRg.Document.Hyperlinks.Add(Anchor:=wrk, Address:=CStr(href))
            If styleAction IsNot Nothing Then
                Debug.WriteLine("[InsertInline] → Anwenden styleAction auf Hyperlink‑Range")
                styleAction(hl.Range)
            End If
            hl.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            mainRg.SetRange(hl.Range.End, hl.Range.End)
        Else
            If styleAction IsNot Nothing Then
                Debug.WriteLine("[InsertInline] → Anwenden styleAction auf Text‑Range")
                styleAction(wrk)
            End If
            wrk.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            mainRg.SetRange(wrk.End, wrk.End)
        End If

    End Sub



End Class
