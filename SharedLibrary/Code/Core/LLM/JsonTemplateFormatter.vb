' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Imports System.Net
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary
    Public Module JsonTemplateFormatter

        ''' <summary>
        ''' Hauptfunktion für JSON-String + Template
        ''' </summary>
        Public Function FormatJsonWithTemplate(json As String, ByVal template As String) As String
            Dim jObj As JObject
            Try
                jObj = JObject.Parse(json)
            Catch ex As Newtonsoft.Json.JsonReaderException
                Return $"[Fehler beim Parsen des JSON: {ex.Message}]"
            End Try
            NormalizeSources(jObj)
            Return FormatJsonWithTemplate(jObj, template)
        End Function

        ''' <summary>
        ''' Hauptfunktion für direkten JObject + Template
        ''' </summary>
        Public Function FormatJsonWithTemplate(jObj As JObject, ByVal template As String, Optional ByVal Mode As Integer = 2) As String
            If String.IsNullOrWhiteSpace(template) Then Return ""

            NormalizeSources(jObj)

            ' Normalize CRLF / Platzhalter für Zeilenumbruch
            template = template _
        .Replace("\N", vbCrLf) _
        .Replace("\n", vbCrLf) _
        .Replace("\R", vbCrLf) _
        .Replace("\r", vbCrLf)
            template = Regex.Replace(template, "<cr>", vbCrLf, RegexOptions.IgnoreCase)

            Dim hasLoop = Regex.IsMatch(template, "\{\%\s*for\s+([^\s\%]+)\s*\%\}", RegexOptions.Singleline)
            Dim hasPh = Regex.IsMatch(template, "\{([^}]+)\}")

            ' === Einfache Fallbehandlung ===
            If Not hasLoop AndAlso Not hasPh Then
                ' Template enthält keine Platzhalter → als einfacher JSONPath behandeln

                Dim ResponseString As String = ""
                Select Case Mode
                    Case 1
                        ' 1) Standard: alle Vorkommen zusammenhängen (Dokumentreihenfolge)
                        ResponseString = FindJsonPropertyCustom(jObj, template, SelectionMode.JoinAll)
                    Case 2
                        ' 2) „Bestes“ einzelnes Vorkommen: längstes nicht-leeres
                        ResponseString = FindJsonPropertyCustom(jObj, template, SelectionMode.LongestNonEmpty)
                    Case 3
                        ' 3) Bewahre exakt dein altes Verhalten (aber ignoriere leere Werte):
                        ResponseString = FindJsonPropertyCustom(jObj, template, SelectionMode.FirstNonEmpty)
                    Case Else
                        ResponseString = FindJsonProperty(jObj, template)

                End Select

                Return ResponseString
            End If


            ' === Schleifen-Blöcke ===
            Dim loopRegex = New Regex("\{\%\s*for\s+([^%\s]+)\s*\%\}(.*?)\{\%\s*endfor\s*\%\}", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            Dim mLoop = loopRegex.Match(template)
            While mLoop.Success
                Dim fullBlock = mLoop.Value
                Dim rawPath = mLoop.Groups(1).Value.Trim()
                Dim innerTpl = mLoop.Groups(2).Value

                Dim path = If(rawPath.StartsWith("$"), rawPath, "$." & rawPath)
                Dim tokens = jObj.SelectTokens(path)
                Dim items = tokens.SelectMany(Function(t)
                                                  If t.Type = JTokenType.Array Then
                                                      Return CType(t, JArray).OfType(Of JObject)()
                                                  ElseIf t.Type = JTokenType.Object Then
                                                      Return {CType(t, JObject)}
                                                  Else
                                                      Return Enumerable.Empty(Of JObject)()
                                                  End If
                                              End Function)

                Dim rendered = items.Select(Function(o) FormatJsonWithTemplate(o, innerTpl)).ToArray()
                template = template.Replace(fullBlock, If(rendered.Any, String.Join(vbCrLf & vbCrLf, rendered), ""))
                mLoop = loopRegex.Match(template)
            End While

            ' === Platzhalter (non-gierig) ===
            Dim phRegex = New Regex("\{(.+?)\}", RegexOptions.Singleline)
            Dim result = template

            For Each mPh As Match In phRegex.Matches(template)
                Dim fullPh = mPh.Value
                Dim content = mPh.Groups(1).Value

                ' HTML- oder No-CR-Flag?
                Dim isHtml As Boolean = False
                Dim isNoCr As Boolean = False

                If content.StartsWith("htmlnocr:", StringComparison.OrdinalIgnoreCase) Then
                    isHtml = True
                    isNoCr = True
                    content = content.Substring("htmlnocr:".Length)
                ElseIf content.StartsWith("html:", StringComparison.OrdinalIgnoreCase) Then
                    isHtml = True
                    content = content.Substring("html:".Length)
                ElseIf content.StartsWith("nocr:", StringComparison.OrdinalIgnoreCase) Then
                    isNoCr = True
                    content = content.Substring("nocr:".Length)
                End If

                ' Nur am ersten "|" trennen
                Dim parts = content.Split(New Char() {"|"c}, 2)
                Dim pathPh = parts(0).Trim()
                Dim remainder = If(parts.Length > 1, parts(1), String.Empty)

                ' Separator-Override (z.B. "/") oder Mapping-Definition (enthält "=")
                Dim sep As String = vbCrLf
                Dim mappings As Dictionary(Of String, String) = Nothing

                If Not String.IsNullOrEmpty(remainder) Then
                    If remainder.Contains("="c) Then
                        mappings = ParseMappings(remainder)
                    Else
                        sep = remainder.Replace("\n", vbCrLf)
                    End If
                End If

                Dim replacement = RenderTokens(jObj, pathPh, sep, isHtml, isNoCr, mappings)
                result = result.Replace(fullPh, replacement)
            Next

            Return result
        End Function


        Public Enum SelectionMode
            FirstNonEmpty = 0     ' Erstes nicht-leeres Vorkommen
            LongestNonEmpty = 1   ' Längstes nicht-leeres Vorkommen (robuster "best"-Heuristik)
            JoinAll = 2           ' Alle zusammenhängen (Dokumentreihenfolge)
        End Enum

        ''' <summary>
        ''' Sucht nach allen Vorkommen eines Property-Namens (case-insensitive) und
        ''' liefert je nach Modus das "beste" oder die Konkatenation.
        ''' </summary>
        ''' <param name="jObj">Wurzel-JSON (Newtonsoft JObject)</param>
        ''' <param name="propertyName">Gesuchter Property-Name (z. B. "text")</param>
        ''' <param name="mode">Auswahlstrategie (Default: JoinAll)</param>
        ''' <param name="separator">Trenner beim Zusammenhängen (Default: Doppel-CRLF)</param>
        Public Function FindJsonPropertyCustom(jObj As JObject,
                                 propertyName As String,
                                 Optional mode As SelectionMode = SelectionMode.JoinAll,
                                 Optional separator As String = vbCrLf & vbCrLf) As String
            If jObj Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(jObj))
            End If
            If String.IsNullOrWhiteSpace(propertyName) Then
                Throw New System.ArgumentException("propertyName darf nicht leer sein.", NameOf(propertyName))
            End If

            Dim matches As System.Collections.Generic.List(Of String) = CollectPropertyValues(jObj, propertyName)

            If matches Is Nothing OrElse matches.Count = 0 Then
                Return String.Empty
            End If

            Select Case mode
                Case SelectionMode.FirstNonEmpty
                    For Each s As String In matches
                        If Not String.IsNullOrWhiteSpace(s) Then
                            Return s
                        End If
                    Next
                    Return matches(0) ' falls alles leer/Whitespace war

                Case SelectionMode.LongestNonEmpty
                    Dim best As String = String.Empty
                    Dim bestLen As Integer = -1
                    For Each s As String In matches
                        Dim candidate As String = If(s, String.Empty).Trim()
                        If candidate.Length > bestLen Then
                            best = candidate
                            bestLen = candidate.Length
                        End If
                    Next
                    Return best

                Case SelectionMode.JoinAll
                    Dim sb As New System.Text.StringBuilder()
                    For Each s As String In matches
                        Dim part As String = If(s, String.Empty).Trim()
                        If part.Length = 0 Then Continue For
                        If sb.Length > 0 Then sb.Append(separator)
                        sb.Append(part)
                    Next
                    Return sb.ToString()

                Case Else
                    ' Fallback sicherheitshalber
                    Return String.Join(separator, matches)

            End Select
        End Function

        ' ---- Hilfsfunktionen ---------------------------------------------------

        Private Function CollectPropertyValues(root As JToken, propertyName As String) As System.Collections.Generic.List(Of String)
            Dim results As New System.Collections.Generic.List(Of String)()
            Dim stack As New System.Collections.Generic.Stack(Of JToken)()
            stack.Push(root)

            While stack.Count > 0
                Dim node As JToken = stack.Pop()

                Select Case node.Type
                    Case JTokenType.Object
                        Dim obj As JObject = CType(node, JObject)
                        ' Properties in umgekehrter Reihenfolge pushen, damit die Iteration links->rechts stabil bleibt
                        Dim props As System.Collections.Generic.IEnumerable(Of JProperty) = obj.Properties()
                        For Each p As JProperty In props.Reverse()
                            stack.Push(p)
                        Next

                    Case JTokenType.Property
                        Dim jp As JProperty = CType(node, JProperty)
                        If jp.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase) Then
                            Dim s As String = ConvertTokenToString(jp.Value)
                            results.Add(s)
                        End If
                        If jp.Value IsNot Nothing Then
                            stack.Push(jp.Value)
                        End If

                    Case JTokenType.Array
                        Dim arr As JArray = CType(node, JArray)
                        For i As Integer = arr.Count - 1 To 0 Step -1
                            stack.Push(arr(i))
                        Next

                    Case Else
                        ' einfache Werte: nichts zu tun
                End Select
            End While

            Return results
        End Function

        Private Function ConvertTokenToString(t As JToken) As String
            If t Is Nothing Then Return String.Empty

            Select Case t.Type
                Case JTokenType.String
                    Return CStr(t)

                Case JTokenType.Integer, JTokenType.Float, JTokenType.Boolean, JTokenType.Null
                    Return t.ToString()

                Case JTokenType.Array, JTokenType.Object
                    ' Für Nicht-Strings geben wir eine kompakte JSON-Serialisierung zurück,
                    ' damit "kein Text verloren geht".
                    Return t.ToString(Newtonsoft.Json.Formatting.None)

                Case Else
                    Return t.ToString()
            End Select
        End Function


        ''' <summary>
        ''' Wandelt ausgewählte Tokens in einen String um, wendet Mapping, HTML→Markdown und No-CR an.
        ''' </summary>
        Private Function RenderTokens(
        jObj As JObject,
        path As String,
        sep As String,
        isHtml As Boolean,
        isNoCr As Boolean,
        mappings As Dictionary(Of String, String)
    ) As String

            Try
                If Not path.StartsWith("$") AndAlso Not path.StartsWith("@") Then
                    path = "$." & path
                End If
                Dim tokens = jObj.SelectTokens(path)
                Dim list As New List(Of String)

                For Each t In tokens
                    Dim raw = t.ToString()
                    ' Mapping anwenden, falls definiert
                    If mappings IsNot Nothing AndAlso mappings.ContainsKey(raw) Then raw = mappings(raw)
                    ' HTML→Markdown, falls gewünscht
                    If isHtml Then raw = HtmlToMarkdownSimple(raw)
                    ' No-CR: alle Zeilenumbrüche durch Leerzeichen
                    'If isNoCr Then raw = Regex.Replace(raw, "[\r\n]+", " ").Trim()
                    If isNoCr Then
                        ' 1) Turn all line-breaks into single spaces
                        raw = Regex.Replace(raw, "[\r\n]+", " ")

                        ' 2) Collapse any run of whitespace into one space
                        raw = Regex.Replace(raw, "\s{2,}", " ")

                        ' 3) Remove common Unicode bullet characters only
                        raw = Regex.Replace(raw, "[\u2022\u2023\u25E6]", String.Empty)

                        ' 4) Trim leading/trailing spaces
                        raw = raw.Trim()
                    End If


                    list.Add(raw)
                Next

                Return If(list.Count = 0, "", String.Join(sep, list))
            Catch ex As System.Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Parst Mapping-Definitionen der Form "key1=Text1;key2=Text2;…"
        ''' </summary>
        Private Function ParseMappings(defs As String) As Dictionary(Of String, String)
            Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each pair In defs.Split(";"c)
                Dim kv = pair.Split(New Char() {"="c}, 2)
                If kv.Length = 2 Then dict(kv(0).Trim()) = kv(1).Trim()
            Next
            Return dict
        End Function

        ''' <summary>
        ''' Einfacher HTML→Markdown-Konverter (inkl. SPAN → *italic*)
        ''' </summary>
        Public Function HtmlToMarkdownSimple(html As String) As String
            Dim s = WebUtility.HtmlDecode(html)
            ' Absätze → zwei Zeilenumbrüche            
            s = Regex.Replace(s, "</?p\s*/?>", vbCrLf & vbCrLf, RegexOptions.IgnoreCase)
            ' Zeilenumbruch-Tags
            s = Regex.Replace(s, "<br\s*/?>", vbCrLf, RegexOptions.IgnoreCase)
            ' Fett/strong → **text**
            s = Regex.Replace(s, "<strong>(.*?)</strong>", "**$1**", RegexOptions.IgnoreCase)
            ' Kursiv/em → *text*
            s = Regex.Replace(s, "<em>(.*?)</em>", "*$1*", RegexOptions.IgnoreCase)
            ' SPAN-Tags → *text*
            s = Regex.Replace(s, "<span\b[^>]*>(.*?)</span>", "*$1*", RegexOptions.IgnoreCase)
            ' Listenpunkte <li> → "- text"
            s = Regex.Replace(s, "<li>(.*?)</li>", "- $1" & vbCrLf, RegexOptions.IgnoreCase)
            ' Fußnoten-Tags <fn>…</fn> → <sup>…</sup>
            s = Regex.Replace(s, "<fn>(.*?)</fn>", "<sup>$1</sup>", RegexOptions.IgnoreCase)
            ' Alle übrigen Tags entfernen
            s = Regex.Replace(s, "<(?!/?sup\b)[^>]+>", String.Empty, RegexOptions.IgnoreCase)
            's = Regex.Replace(s, "<[^>]+>", String.Empty)
            ' Mehrfache Zeilenumbrüche aufräumen
            s = Regex.Replace(s, "(" & vbCrLf & "){3,}", vbCrLf & vbCrLf)
            Return s.Trim()
        End Function

        Private Sub NormalizeSources(jObj As JObject)
            Dim srcToken = jObj.SelectToken("sources")
            If srcToken IsNot Nothing AndAlso srcToken.Type = JTokenType.Array Then
                Dim newArray As New JArray()
                For Each item In CType(srcToken, JArray)
                    If item.Type = JTokenType.Array AndAlso item.Count >= 3 Then
                        Dim objStr = item(2).ToString()
                        Try
                            Dim o = JObject.Parse(objStr)
                            newArray.Add(o)
                        Catch ex As System.Exception
                            ' Ungültiges JSON überspringen
                        End Try
                    ElseIf item.Type = JTokenType.Object Then
                        newArray.Add(item)
                    End If
                Next
                jObj("sources") = newArray
            End If
        End Sub

    End Module


End Namespace