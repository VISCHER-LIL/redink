' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.FindClause.vb
' Purpose: Provides the FindClause and AddClause workflows for the Word add-in, handling
'          clause retrieval via LLM and maintenance of clause library files.
'
' Architecture:
'  - ClauseLibrary parsing: Scans global/local folders for AN2-lib-*.txt files, splits them into
'    segments, and resolves prompt overrides at file or segment scope.
'  - UI parameter collection: Uses SharedMethods input forms for library selection, search text,
'    and optional library editing entry points.
'  - LLM interaction: Builds system/user prompts (with selected text, explicit queries, and JSON
'    segments), optionally swaps model configuration, and parses the LLM response into Markdown.
'  - Markdown rendering: Cleans LLM JSON, tolerates malformed payloads, and produces pane-ready
'    Markdown with metadata.
'  - Library maintenance: AddClause appends text into the chosen segment while preserving the
'    original JSON layout (wrapper object, array, or standalone objects) and enforces duplicate
'    checks.
'  - Helper utilities: Include directory enumeration, INI path expansion, JSON cleanup/fallback
'    parsing, segment range tracking, and duplicate detection across clause formats.
' =============================================================================


Option Explicit On
Option Strict On

Imports System.Collections
Imports System.Data
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Encapsulates clause discovery and clause library maintenance workflows for Red Ink for Word.
''' </summary>
Partial Public Class ThisAddIn

    ''' <summary>
    ''' Represents a clause library segment loaded from a library file, including resolved prompt overrides.
    ''' </summary>
    Private Class ClauseLibrary
        Public Property Title As String
        Public Property SourcePath As String
        Public Property IsLocal As Boolean
        Public Property RawJson As String                  ' JSON for that library segment
        Public Property EffectivePrompt As String          ' FindClause prompt override (file or segment)
        Public Property EffectiveMergePrompt As String     ' MergePrompt override (file or segment)  << NEW
    End Class

    ''' <summary>
    ''' Executes the Clause Finder workflow: gathers user input, calls the LLM, and displays matching clauses.
    ''' </summary>
    Public Async Sub FindClause()
        Try
            If INILoadFail() Then Return

            ' 1) Expand / normalize paths
            Dim pathGlobal As String = ExpandEnvironmentVariables(INI_FindClausePath)
            If Not String.IsNullOrEmpty(pathGlobal) AndAlso Not pathGlobal.EndsWith("\", StringComparison.Ordinal) Then pathGlobal &= "\"
            Dim pathLocal As String = ExpandEnvironmentVariables(INI_FindClausePathLocal)
            If Not String.IsNullOrEmpty(pathLocal) AndAlso Not pathLocal.EndsWith("\", StringComparison.Ordinal) Then pathLocal &= "\"

            Dim hasGlobal As Boolean = Not String.IsNullOrWhiteSpace(pathGlobal) AndAlso IO.Directory.Exists(pathGlobal)
            Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(pathLocal) AndAlso IO.Directory.Exists(pathLocal)

            If Not hasGlobal AndAlso Not hasLocal Then
                ShowCustomMessageBox("No Clause Library paths are configured or accessible. Configure 'ClauseFindPath' or 'ClauseFindPathLocal'.")
                Return
            End If

            ' 2) Acquire current selection context
            Dim app As Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
                ShowCustomMessageBox("No active document.")
                Return
            End If

            Dim doc As Word.Document = app.ActiveDocument
            Dim sel As Word.Selection = app.Selection

            Dim selectedText As String = ""
            If sel.Range IsNot Nothing AndAlso sel.Range.Text IsNot Nothing Then
                selectedText = sel.Range.Text
                If selectedText IsNot Nothing Then selectedText = selectedText.Trim()
            End If

            ' 3) Load all libraries (may be multiple segments per file)
            Dim allLibs As List(Of ClauseLibrary) = LoadClauseLibraries(pathGlobal, pathLocal)
            If allLibs Is Nothing OrElse allLibs.Count = 0 Then
                ShowCustomMessageBox($"No clause libraries found. Place files named '{AN2}-lib-*.txt' in the configured path(s).")
                Return
            End If

            ' 4) Build dropdown entries
            allLibs.Sort(Function(a, b) String.Compare(a.Title, b.Title, StringComparison.OrdinalIgnoreCase))
            Dim displayMap As New Dictionary(Of String, ClauseLibrary)(StringComparer.OrdinalIgnoreCase)
            Dim options As New List(Of String)
            For Each cl In allLibs
                Dim disp As String = cl.Title
                If cl.IsLocal Then disp &= " (local)"
                disp = MakeUniqueDisplay(disp, displayMap.Keys) ' uniqueness helper
                displayMap(disp) = cl
                options.Add(disp)
            Next

            ' 5) Parameters dialog
            Dim defaultDisplay As String = If(options.Count > 0, options(0), "")
            Dim defaultUseSelected As Boolean = Not String.IsNullOrWhiteSpace(selectedText)
            ' OtherPrompt (global) will hold the search query
            Dim p0 As New SLib.InputParameter("Clause Library", defaultDisplay) With {.Options = New List(Of String)(options)}
            Dim p1 As New SLib.InputParameter("Search for", "") ' text to search (query)
            Dim paramArr() As SLib.InputParameter
            If selectedText <> "" Then
                Dim p2 As New SLib.InputParameter("Use selected text", defaultUseSelected)
                paramArr = {p0, p1, p2}
            Else
                paramArr = {p0, p1}
            End If

            ' Optional extra button: “Edit Library File…”
            Dim extraText As String = Nothing
            Dim extraAction As System.Action = Nothing
            Dim closeAfterExtra As Boolean = False

            If hasGlobal OrElse hasLocal Then
                extraText = "Edit Library File…"
                extraAction =
                    Sub()
                        Try

                            ' Build list of available library files, same as in AddClause (global + local)
                            Dim displayToPath As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                            Dim editoptions As New List(Of String)

                            Dim dcPaths As New List(Of (p As String, isLocal As Boolean))
                            If hasGlobal AndAlso Not String.IsNullOrWhiteSpace(pathGlobal) Then
                                dcPaths.Add((pathGlobal, False))
                            End If
                            If hasLocal AndAlso Not String.IsNullOrWhiteSpace(pathLocal) Then
                                dcPaths.Add((pathLocal, True))
                            End If

                            For Each tuple In dcPaths
                                Dim basePath = tuple.p
                                Dim isLocal = tuple.isLocal
                                If IO.Directory.Exists(basePath) Then
                                    Dim files = IO.Directory.GetFiles(basePath, $"{AN2}-lib-*.txt", IO.SearchOption.TopDirectoryOnly)
                                    For Each f In files
                                        Dim disp As String = IO.Path.GetFileName(f)
                                        If isLocal Then disp &= " (local)"
                                        If Not displayToPath.ContainsKey(disp) Then
                                            displayToPath.Add(disp, f)
                                            editoptions.Add(disp)
                                        End If
                                    Next
                                End If
                            Next

                            If editoptions.Count = 0 Then
                                SLib.ShowCustomMessageBox($"No FindClause library files ({AN2}-lib-*.txt) found in the configured paths.")
                                Exit Sub
                            End If

                            ' Let user pick one (same helper as in AddClause)
                            Dim selfile As String = SLib.ShowSelectionForm("Select a library file to view or edit:", $"{AN} FindClause library files", editoptions)
                            If String.IsNullOrWhiteSpace(selfile) Then Exit Sub

                            Dim chosenPath As String = Nothing
                            If displayToPath.TryGetValue(selfile, chosenPath) AndAlso Not String.IsNullOrWhiteSpace(chosenPath) Then
                                SLib.ShowTextFileEditor(chosenPath, $"{AN} FindClause library file '{chosenPath}':", True, _context)
                                SLib.ShowCustomMessageBox("Any changes to the library will only be active the next time this feature is called up.")
                            End If

                        Catch ex As Exception
                            SLib.ShowCustomMessageBox("Error while opening a library file:" & vbCrLf & ex.Message)
                            Exit Sub
                        End Try

                    End Sub
            End If

            If ShowCustomVariableInputForm("Please set the Clause Finder parameters:", AN & " FindClause", paramArr, extraButtonText:=extraText,
                                                                                                                            extraButtonAction:=extraAction,
                                                                                                                            CloseAfterExtra:=closeAfterExtra) = False Then
                Return
            End If

            OtherPrompt = ""
            Dim chosenDisplay As String = CStr(paramArr(0).Value)
            OtherPrompt = CStr(paramArr(1).Value)          ' store query globally as required
            Dim useSelected As Boolean = False
            If paramArr.Length > 2 Then useSelected = CBool(paramArr(2).Value)

            If String.IsNullOrWhiteSpace(OtherPrompt) AndAlso Not useSelected Then
                ShowCustomMessageBox("You have not provided a search term - will abort.")
                Return
            End If

            Dim chosenLib As ClauseLibrary = Nothing
            If Not displayMap.TryGetValue(chosenDisplay, chosenLib) OrElse chosenLib Is Nothing Then
                ShowCustomMessageBox("Selected library could not be resolved - will abort.")
                Return
            End If

            ' 6) Build prompts
            Dim systemPrompt As String = InterpolateAtRuntime(If(String.IsNullOrWhiteSpace(chosenLib.EffectivePrompt), SP_FindClause, chosenLib.EffectivePrompt))

            Dim userPromptBuilder As New System.Text.StringBuilder()
            If useSelected AndAlso Not String.IsNullOrWhiteSpace(selectedText) Then
                userPromptBuilder.AppendLine("<TEXTFORSEARCH>")
                userPromptBuilder.AppendLine(selectedText)
                userPromptBuilder.AppendLine("</TEXTFORSEARCH>")
            End If
            If Not String.IsNullOrWhiteSpace(OtherPrompt) Then
                ' Provide search query explicitly
                userPromptBuilder.AppendLine("<SEARCHQUERY>")
                userPromptBuilder.AppendLine(OtherPrompt)
                userPromptBuilder.AppendLine("</SEARCHQUERY>")
            End If
            ' Provide library JSON (raw) between tags
            userPromptBuilder.AppendLine("<LIBRARY>")
            userPromptBuilder.AppendLine(chosenLib.RawJson.Trim())
            userPromptBuilder.AppendLine("</LIBRARY>")

            Dim userPrompt As String = userPromptBuilder.ToString()

            ' 7) Call LLM

            Dim UseSecondAPI As Boolean = False

            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                If Not GetSpecialTaskModel(_context, INI_AlternateModelPath, "FindClause") Then
                    originalConfigLoaded = False
                Else
                    UseSecondAPI = True
                End If
            End If

            Dim llmResponse As String = Await LLM(systemPrompt, userPrompt, "", "", 0, UseSecondAPI)

            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

            If String.IsNullOrWhiteSpace(llmResponse) Then
                ShowCustomMessageBox("No response received from the model.")
                Return
            End If
            llmResponse = llmResponse.Trim()

            ' 8) Parse response JSON → build Markdown
            Dim markdownResult As String = BuildMarkdownFromClauseResponse(llmResponse)

            ' 9) Show in pane (Markdown conversion handled by ShowPaneAsync with insertMarkdown:=True)
            Dim paneHeader As String = $"The following clauses were found:"
            Dim paneFooter As String = "Select the clause you want to use and click on merge, copy or insert to do so."
            Dim finalToShow As String = markdownResult

            ' Decide which merge prompt to cache (segment > file-level > default)
            Dim mergePromptToUse As String = If(String.IsNullOrWhiteSpace(chosenLib.EffectiveMergePrompt), SP_MergePrompt, chosenLib.EffectiveMergePrompt)

            If _uiContext IsNot Nothing Then
                Dim localMergePrompt = mergePromptToUse ' capture for closure safety
                _uiContext.Post(Sub(s)
                                    SP_MergePrompt_Cached = localMergePrompt
                                    ShowPaneAsync(paneHeader, finalToShow, paneFooter, AN, noRTF:=False, insertMarkdown:=True)
                                End Sub, Nothing)
            Else
                SP_MergePrompt_Cached = mergePromptToUse
                ShowPaneAsync(paneHeader, finalToShow, paneFooter, AN, noRTF:=False, insertMarkdown:=True)
            End If

        Catch ex As System.Exception
#If DEBUG Then
            System.Diagnostics.Debug.WriteLine("Error: " & ex.Message)
            System.Diagnostics.Debug.WriteLine("Stacktrace: " & ex.StackTrace)

            System.Diagnostics.Debugger.Break()
#End If
            ShowCustomMessageBox("Error in FindClause: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Loads all clause library segments from both global and local paths.
    ''' </summary>
    ''' <param name="pathGlobal">Expanded path to the global library folder.</param>
    ''' <param name="pathLocal">Expanded path to the local library folder.</param>
    ''' <returns>List of clause segments discovered in the configured paths.</returns>
    Private Function LoadClauseLibraries(pathGlobal As String, pathLocal As String) As List(Of ClauseLibrary)
        Dim list As New List(Of ClauseLibrary)()
        Dim candidates As New List(Of Tuple(Of String, Boolean))()

        If Not String.IsNullOrWhiteSpace(pathGlobal) AndAlso IO.Directory.Exists(pathGlobal) Then
            For Each f In EnumerateClauseLibraryFiles(pathGlobal)
                candidates.Add(Tuple.Create(f, False))
            Next
        End If
        If Not String.IsNullOrWhiteSpace(pathLocal) AndAlso IO.Directory.Exists(pathLocal) Then
            For Each f In EnumerateClauseLibraryFiles(pathLocal)
                candidates.Add(Tuple.Create(f, True))
            Next
        End If

        For Each t In candidates
            list.AddRange(ParseClauseLibraryFile(t.Item1, t.Item2))
        Next

        Return list
    End Function

    ''' <summary>
    ''' Enumerates clause library files inside the provided folder.
    ''' </summary>
    ''' <param name="folder">Target directory to scan.</param>
    ''' <returns>Enumerable of absolute file paths that match the AN2-lib-*.txt pattern.</returns>
    Private Function EnumerateClauseLibraryFiles(folder As String) As IEnumerable(Of String)
        Dim matches As New List(Of String)
        Try
            For Each f In IO.Directory.EnumerateFiles(folder, $"{AN2}-lib-*.txt", IO.SearchOption.TopDirectoryOnly)
                matches.Add(f)
            Next
        Catch ex As Exception
            ShowCustomMessageBox("Failed to enumerate library files in '" & folder & "': " & ex.Message)
        End Try
        Return matches
    End Function

    ''' <summary>
    ''' Parses a clause library file into ClauseLibrary instances, capturing prompt overrides per segment.
    ''' </summary>
    ''' <param name="filePath">Library file to read.</param>
    ''' <param name="isLocal">Indicates whether the file resides in the local directory.</param>
    ''' <returns>List of segments discovered within the file.</returns>
    Private Function ParseClauseLibraryFile(filePath As String, isLocal As Boolean) As List(Of ClauseLibrary)
        Dim libs As New List(Of ClauseLibrary)()
        Try
            Dim fileDefaultFindPrompt As String = Nothing
            Dim fileDefaultMergePrompt As String = Nothing
            Dim currentTitle As String = Nothing
            Dim segFindPrompt As String = Nothing
            Dim segMergePrompt As String = Nothing
            Dim jsonBuilder As New System.Text.StringBuilder()

            Dim FlushCurrent As System.Action =
        Sub()
            Dim raw = jsonBuilder.ToString().Trim()
            If currentTitle IsNot Nothing AndAlso raw.Length > 0 Then
                Dim effFind As String = If(segFindPrompt, fileDefaultFindPrompt)
                Dim effMerge As String = If(segMergePrompt, fileDefaultMergePrompt)
                libs.Add(New ClauseLibrary With {
                    .Title = currentTitle,
                    .SourcePath = filePath,
                    .IsLocal = isLocal,
                    .RawJson = raw,
                    .EffectivePrompt = effFind,
                    .EffectiveMergePrompt = effMerge
                })
            End If
            jsonBuilder.Clear()
            segFindPrompt = Nothing
            segMergePrompt = Nothing
        End Sub

            For Each rawLine In IO.File.ReadLines(filePath)
                If rawLine Is Nothing Then Continue For
                Dim line As String = rawLine.Trim()

                If line.StartsWith(";", StringComparison.Ordinal) Then Continue For

                ' FindClause override?
                Dim findVal As String = Nothing
                If TryParseFindClauseLine(line, findVal) Then
                    If currentTitle Is Nothing Then
                        fileDefaultFindPrompt = findVal
                    Else
                        segFindPrompt = findVal
                    End If
                    Continue For
                End If

                ' MergePrompt override?
                Dim mergeVal As String = Nothing
                If TryParseMergePromptLine(line, mergeVal) Then
                    If currentTitle Is Nothing Then
                        fileDefaultMergePrompt = mergeVal
                    Else
                        segMergePrompt = mergeVal
                    End If
                    Continue For
                End If

                ' New segment start
                If line.StartsWith("[", StringComparison.Ordinal) AndAlso line.EndsWith("]", StringComparison.Ordinal) Then
                    FlushCurrent()
                    currentTitle = line.Substring(1, line.Length - 2).Trim()
                    Continue For
                End If

                If currentTitle IsNot Nothing Then
                    jsonBuilder.AppendLine(rawLine)
                End If
            Next
            FlushCurrent()

        Catch ex As Exception
            ShowCustomMessageBox("Failed to parse library file '" & filePath & "': " & ex.Message)
        End Try
        Return libs
    End Function

    ''' <summary>
    ''' Attempts to parse a FindClause prompt override line.
    ''' </summary>
    ''' <param name="line">Line content to evaluate.</param>
    ''' <param name="valueOut">Outputs the parsed prompt text.</param>
    ''' <returns>True when the line contains an SP_FindClause assignment.</returns>
    Private Function TryParseFindClauseLine(line As String, ByRef valueOut As String) As Boolean
        valueOut = Nothing
        If line Is Nothing Then Return False
        Dim m = System.Text.RegularExpressions.Regex.Match(line, "^\s*SP_FindClause\s*=\s*(.*)$", RegexOptions.IgnoreCase)
        If m IsNot Nothing AndAlso m.Success Then
            valueOut = m.Groups(1).Value.Trim()
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' Attempts to parse a MergePrompt override line.
    ''' </summary>
    ''' <param name="line">Line content to evaluate.</param>
    ''' <param name="valueOut">Outputs the parsed prompt text.</param>
    ''' <returns>True when the line contains an SP_MergePrompt assignment.</returns>
    Private Function TryParseMergePromptLine(line As String, ByRef valueOut As String) As Boolean
        valueOut = Nothing
        If line Is Nothing Then Return False
        Dim m = System.Text.RegularExpressions.Regex.Match(line, "^\s*SP_MergePrompt\s*=\s*(.*)$", RegexOptions.IgnoreCase)
        If m IsNot Nothing AndAlso m.Success Then
            valueOut = m.Groups(1).Value.Trim()
            Return True
        End If
        Return False
    End Function

    ' ================== Revised BuildMarkdownFromClauseResponse ==================
    ' Expects the LLM to return:
    ' {
    '   "records":[
    '     {
    '       "clause":"<verbatim>",
    '       "title":"<optional>",
    '       "id":"<optional>",
    '       "score":0.87,
    '       "reason":"<optional short rationale>"
    '     }, ...
    '   ]
    ' }
    ' Falls back gracefully if the structure deviates.
    ' Replace the previous BuildMarkdownFromClauseResponse with the improved, more robust version below.
    ' Add the two new helper functions (CleanAndExtractJson, FallbackExtractRecords) anywhere in the class.

    ''' <summary>
    ''' Cleans LLM output and extracts the most probable JSON payload for downstream parsing.
    ''' </summary>
    ''' <param name="raw">Raw LLM response.</param>
    ''' <returns>String containing the cleaned JSON region or an empty string.</returns>
    Private Function CleanAndExtractJson(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return ""
        Dim s = raw.Trim()

        ' Strip leading / trailing markdown fences ```...```
        ' Accept variants: ```json, ```JSON, ```
        If s.StartsWith("```", StringComparison.Ordinal) Then
            Dim idx = s.IndexOf(vbLf)
            If idx > -1 Then
                s = s.Substring(idx + 1).TrimStart()
            End If
        End If
        If s.EndsWith("```", StringComparison.Ordinal) Then
            Dim lastFence = s.LastIndexOf("```", StringComparison.Ordinal)
            If lastFence >= 0 Then
                s = s.Substring(0, lastFence).TrimEnd()
            End If
        End If

        ' If model wrapped JSON in prose, try to isolate from first { to last }
        Dim firstBrace = s.IndexOf("{"c)
        Dim lastBrace = s.LastIndexOf("}"c)
        If firstBrace >= 0 AndAlso lastBrace > firstBrace Then
            s = s.Substring(firstBrace, lastBrace - firstBrace + 1).Trim()
        End If

        Return s
    End Function

    ''' <summary>
    ''' Extracts record-like dictionaries from malformed JSON as a last resort.
    ''' </summary>
    ''' <param name="raw">Raw or partially cleaned response string.</param>
    ''' <returns>List of dictionaries resembling clause records.</returns>
    Private Function FallbackExtractRecords(raw As String) As List(Of Dictionary(Of String, String))
        Dim list As New List(Of Dictionary(Of String, String))()
        If String.IsNullOrWhiteSpace(raw) Then Return list

        ' Try to isolate the records array content
        Dim m = System.Text.RegularExpressions.Regex.Match(raw, """records""\s*:\s*\[(.*)\]\s*\}", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
        If Not m.Success Then Return list

        Dim arrayContent = m.Groups(1).Value

        ' Split naively on "},{" boundaries (add back braces)
        Dim parts = System.Text.RegularExpressions.Regex.Split(arrayContent, "\},\s*\{", RegexOptions.Singleline)
        For Each partRaw In parts
            Dim part = partRaw.Trim()
            If part = "" Then Continue For
            If Not part.StartsWith("{") Then part = "{" & part
            If Not part.EndsWith("}") Then part &= "}"

            Dim d As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            ' pull clause (multiline)
            Dim mClause = System.Text.RegularExpressions.Regex.Match(part, """clause""\s*:\s*""(.*?)""\s*(,|\})", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            If mClause.Success Then
                Dim c = mClause.Groups(1).Value
                c = c.Replace("\r", "").Replace("\n", vbCrLf)
                d("clause") = c
            End If

            Dim mTitle = System.Text.RegularExpressions.Regex.Match(part, """title""\s*:\s*""(.*?)""\s*(,|\})", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            If mTitle.Success Then d("title") = mTitle.Groups(1).Value

            Dim mId = System.Text.RegularExpressions.Regex.Match(part, """id""\s*:\s*""(.*?)""\s*(,|\})", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            If mId.Success Then d("id") = mId.Groups(1).Value

            Dim mReason = System.Text.RegularExpressions.Regex.Match(part, """reason""\s*:\s*""(.*?)""\s*(,|\})", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            If mReason.Success Then d("reason") = mReason.Groups(1).Value

            Dim mScore = System.Text.RegularExpressions.Regex.Match(part, """score""\s*:\s*([0-9]+(\.[0-9]+)?)", RegexOptions.IgnoreCase)
            If mScore.Success Then d("score") = mScore.Groups(1).Value

            If d.ContainsKey("clause") Then list.Add(d)
        Next

        Return list
    End Function

    ''' <summary>
    ''' Converts the LLM JSON response into Markdown, using fallback parsing if required.
    ''' </summary>
    ''' <param name="responseJson">Response string returned by the model.</param>
    ''' <returns>Markdown representation of the discovered clauses.</returns>
    Private Function BuildMarkdownFromClauseResponse(responseJson As String) As String
        Dim sb As New System.Text.StringBuilder()

        If String.IsNullOrWhiteSpace(responseJson) Then
            sb.AppendLine("_No response (empty)._")
            Return sb.ToString().Trim()
        End If

        ' 1. Clean / extract JSON payload
        Dim cleaned = CleanAndExtractJson(responseJson)

        ' 2. First parse attempt (strict)
        Dim topToken As JToken = Nothing
        Dim parsedOk As Boolean = False
        If Not String.IsNullOrWhiteSpace(cleaned) Then
            Try
                ' Repair common LLM issue: unescaped inner quotes starting a value with ""
                ' We ONLY patch inside "clause" value boundaries where a leading "" occurs.
                cleaned = System.Text.RegularExpressions.Regex.Replace(
                cleaned,
                "(?<=\bclause\b\s*:\s*)""""",
                """\\""",
                RegexOptions.IgnoreCase)

                topToken = JToken.Parse(cleaned)
                parsedOk = True
            Catch ex As Exception
                parsedOk = False
            End Try
        End If

        Dim recs As New List(Of JObject)

        If parsedOk Then
            ' 3. Normal extraction
            Dim recordsToken As JToken = Nothing

            If topToken.Type = JTokenType.Object Then
                Dim obj = CType(topToken, JObject)
                Dim prop = obj.Properties().FirstOrDefault(Function(p) p.Name.Equals("records", StringComparison.OrdinalIgnoreCase))
                If prop IsNot Nothing AndAlso prop.Value.Type = JTokenType.Array Then
                    recordsToken = prop.Value
                ElseIf obj.Properties().Any(Function(p) p.Name.Equals("clause", StringComparison.OrdinalIgnoreCase)) Then
                    ' Single record object
                    recordsToken = New JArray(obj)
                End If
            ElseIf topToken.Type = JTokenType.Array Then
                recordsToken = topToken
            End If

            If recordsToken IsNot Nothing Then
                For Each r In recordsToken
                    If r.Type = JTokenType.Object Then recs.Add(CType(r, JObject))
                Next
            End If
        End If

        ' 4. Fallback if strict parse failed or no records found
        ' === FIX: Replace ONLY the fallback block inside BuildMarkdownFromClauseResponse that used GetValueOrDefault ===
        ' Locate the section:  If recs.Count = 0 Then ... If Not parsedOk Then
        ' Replace that entire inner fallback rendering with the code below.

        If recs.Count = 0 Then
            If Not parsedOk Then
                Dim fallback = FallbackExtractRecords(cleaned)
                If fallback.Count > 0 Then
                    Dim idx = 1
                    For Each d In fallback
                        Dim clauseText As String = Nothing
                        If Not d.TryGetValue("clause", clauseText) OrElse String.IsNullOrWhiteSpace(clauseText) Then Continue For

                        Dim titleTxt As String = Nothing
                        If Not d.TryGetValue("title", titleTxt) OrElse String.IsNullOrWhiteSpace(titleTxt) Then
                            If Not d.TryGetValue("id", titleTxt) OrElse String.IsNullOrWhiteSpace(titleTxt) Then
                                titleTxt = "Clause " & idx.ToString()
                            End If
                        End If

                        sb.AppendLine("### " & titleTxt.Trim())
                        sb.AppendLine()
                        sb.AppendLine(clauseText.Trim())
                        sb.AppendLine()

                        Dim meta As New List(Of String)
                        Dim idVal As String = Nothing
                        If d.TryGetValue("id", idVal) AndAlso Not String.IsNullOrWhiteSpace(idVal) AndAlso Not idVal.Equals(titleTxt, StringComparison.OrdinalIgnoreCase) Then
                            meta.Add("ID: " & idVal)
                        End If
                        Dim scoreVal As String = Nothing
                        If d.TryGetValue("score", scoreVal) AndAlso Not String.IsNullOrWhiteSpace(scoreVal) Then
                            meta.Add("Score: " & scoreVal)
                        End If
                        Dim reasonVal As String = Nothing
                        If d.TryGetValue("reason", reasonVal) AndAlso Not String.IsNullOrWhiteSpace(reasonVal) Then
                            meta.Add("Reason: " & reasonVal)
                        End If

                        If meta.Count > 0 Then
                            sb.AppendLine("_" & String.Join(" • ", meta) & "_")
                            sb.AppendLine()
                        End If
                        idx += 1
                    Next
                    Return sb.ToString().Trim()
                Else
                    sb.AppendLine("_Could not parse JSON (even after fallback). Raw cleaned content:_")
                    sb.AppendLine()
                    sb.AppendLine("```json")
                    sb.AppendLine(If(cleaned, ""))
                    sb.AppendLine("```")
                    Return sb.ToString().Trim()
                End If
            Else
                sb.AppendLine("_No matching clauses returned._")
                Return sb.ToString().Trim()
            End If
        End If

        ' 5. Normal markdown rendering for parsed records
        Dim counter As Integer = 1
        For Each ro In recs
            Dim clauseProp = ro.Properties().FirstOrDefault(Function(p) p.Name.Equals("clause", StringComparison.OrdinalIgnoreCase))
            If clauseProp Is Nothing Then Continue For
            Dim clauseText = clauseProp.Value.ToString()
            If String.IsNullOrWhiteSpace(clauseText) Then Continue For

            Dim titleTxt As String = GetFirstString(ro, {"title"})
            If String.IsNullOrWhiteSpace(titleTxt) Then titleTxt = GetFirstString(ro, {"id"})
            If String.IsNullOrWhiteSpace(titleTxt) Then titleTxt = "Clause " & counter

            Dim idTxt = GetFirstString(ro, {"id"})
            Dim reasonTxt = GetFirstString(ro, {"reason", "rationale"})
            Dim scoreTxt As String = Nothing
            Dim scoreProp = ro.Properties().FirstOrDefault(Function(p) p.Name.Equals("score", StringComparison.OrdinalIgnoreCase))
            If scoreProp IsNot Nothing AndAlso scoreProp.Value.Type <> JTokenType.Object AndAlso scoreProp.Value.Type <> JTokenType.Array Then
                scoreTxt = scoreProp.Value.ToString()
            End If

            sb.AppendLine("## " & titleTxt.Trim())
            sb.AppendLine()
            sb.AppendLine(clauseText.Trim())
            sb.AppendLine()

            Dim meta As New List(Of String)
            If Not String.IsNullOrWhiteSpace(idTxt) AndAlso Not idTxt.Equals(titleTxt, StringComparison.OrdinalIgnoreCase) Then meta.Add("ID: " & idTxt)
            If Not String.IsNullOrWhiteSpace(scoreTxt) Then meta.Add("Score: " & scoreTxt)
            If Not String.IsNullOrWhiteSpace(reasonTxt) Then meta.Add("Reason: " & reasonTxt)
            If meta.Count > 0 Then
                sb.AppendLine("_" & String.Join(" • ", meta) & "_")
                sb.AppendLine()
            End If

            counter += 1
        Next

        If counter = 1 Then
            sb.AppendLine("_No usable clause records were found (all malformed or missing 'clause')._")
        End If

        Return sb.ToString().Trim()
    End Function

    ''' <summary>
    ''' Retrieves the first available string value from a JObject using the provided key order.
    ''' </summary>
    ''' <param name="obj">JObject to inspect.</param>
    ''' <param name="keys">Key preference order.</param>
    ''' <returns>The first non-empty string value or Nothing.</returns>
    Private Function GetFirstString(obj As JObject, keys As IEnumerable(Of String)) As String
        For Each k In keys
            Dim p = obj.Properties().FirstOrDefault(Function(pr) pr.Name.Equals(k, StringComparison.OrdinalIgnoreCase))
            If p IsNot Nothing AndAlso p.Value IsNot Nothing Then
                If p.Value.Type = JTokenType.String Then
                    Dim v = p.Value.ToString().Trim()
                    If v.Length > 0 Then Return v
                Else
                    ' If it's not a string but scalar, use ToString
                    If p.Value.Type <> JTokenType.Object AndAlso p.Value.Type <> JTokenType.Array Then
                        Dim v = p.Value.ToString().Trim()
                        If v.Length > 0 Then Return v
                    End If
                End If
            End If
        Next
        Return Nothing
    End Function

    ' ====================== Add Clause to Library (Segment + Lenient JSON) ======================
    ' Lets user pick a specific library SEGMENT (like FindClause) and appends the selected text
    ' ONLY to that segment. Other segments / prompt overrides remain untouched.
    ' Supports three storage styles inside a segment:
    '   (A) { "Records":[ { ... }, ... ] }
    '   (B) [ { ... }, { ... } ]
    '   (C) Sequence of standalone objects:  { ... }\r\n{ ... }\r\n{ ... }
    ' The field name used for clause text is determined from the LAST existing object’s first string property;
    ' if none exists, it falls back to "Text".  Duplicate detection uses that dynamic field.


    ''' <summary>
    ''' Adds the current Word selection (or optionally edited text) to a chosen clause library segment.
    ''' When no text is selected, allows direct editing of a library file. Supports three JSON storage
    ''' formats and preserves the original structure when appending new records.
    ''' </summary>
    ''' <remarks>
    ''' This method performs the following steps:
    ''' 1. Acquires the current Word selection.
    ''' 2. Loads all available clause library segments from configured paths.
    ''' 3. Prompts user to choose a target segment and optional text cleaning.
    ''' 4. Optionally invokes an LLM to clean/anonymize the selected text.
    ''' 5. Parses the target segment's JSON content.
    ''' 6. Appends the new clause while preserving the original JSON layout (wrapper object, array, or standalone objects).
    ''' 7. Performs duplicate detection and writes changes back to the library file with retry logic on file locks.
    ''' </remarks>
    Public Async Sub AddClause()
        Try
            If INILoadFail() Then Return

            ' 1) Get selection
            Dim app As Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
                ShowCustomMessageBox("No active document.")
                Return
            End If
            Dim sel As Word.Selection = app.Selection
            Dim selectedText As String = ""
            If sel IsNot Nothing AndAlso sel.Range IsNot Nothing AndAlso sel.Range.Text IsNot Nothing Then
                selectedText = sel.Range.Text
            End If
            selectedText = If(selectedText, "").Trim()

            If String.IsNullOrWhiteSpace(selectedText) Then
                Dim answer As Integer = ShowCustomYesNoBox("No text is selected to be added to a library. Do you want to manually edit a FindClause library file?", "Yes", "No, abort", AN & " AddClause")
                If answer <> 1 Then Return
            End If

            ' 2) Load all segment libraries (same as FindClause logic)
            Dim pathGlobal As String = ExpandEnvironmentVariables(INI_FindClausePath)
            If Not String.IsNullOrEmpty(pathGlobal) AndAlso Not pathGlobal.EndsWith("\", StringComparison.Ordinal) Then pathGlobal &= "\"
            Dim pathLocal As String = ExpandEnvironmentVariables(INI_FindClausePathLocal)
            If Not String.IsNullOrEmpty(pathLocal) AndAlso Not pathLocal.EndsWith("\", StringComparison.Ordinal) Then pathLocal &= "\"

            Dim allLibs As List(Of ClauseLibrary)

            If Not String.IsNullOrWhiteSpace(selectedText) Then

                allLibs = LoadClauseLibraries(pathGlobal, pathLocal)
                If allLibs Is Nothing OrElse allLibs.Count = 0 Then
                    ShowCustomMessageBox($"No clause library segments found. ({AN2}-lib-*.txt)")
                    Return
                End If

            Else

                Try
                    Dim displayToPath As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    Dim options As New List(Of String)

                    Dim dcPaths As New List(Of (p As String, isLocal As Boolean))
                    If Not String.IsNullOrWhiteSpace(pathGlobal) Then
                        dcPaths.Add((pathGlobal, False))
                    End If
                    If Not String.IsNullOrWhiteSpace(pathLocal) Then
                        dcPaths.Add((pathLocal, True))
                    End If

                    For Each tuple In dcPaths
                        Dim basePath = tuple.p
                        Dim isLocal = tuple.isLocal
                        If IO.Directory.Exists(basePath) Then
                            Dim files = IO.Directory.GetFiles(basePath, $"{AN2}-lib-*.txt", IO.SearchOption.TopDirectoryOnly)
                            For Each f In files
                                Dim disp As String = IO.Path.GetFileName(f)
                                If isLocal Then disp &= " (local)"
                                If Not displayToPath.ContainsKey(disp) Then
                                    displayToPath.Add(disp, f)
                                    options.Add(disp)
                                End If
                            Next
                        End If
                    Next

                    If options.Count = 0 Then
                        SLib.ShowCustomMessageBox($"No FindClause library files ({AN2}-lib-*.txt) found in the configured paths.")
                        Exit Sub
                    End If

                    ' Let user pick one
                    Dim selfile As String = SLib.ShowSelectionForm("Select a library file to view or edit:", $"{AN} FindClause library files", options)
                    If String.IsNullOrWhiteSpace(selfile) Then Exit Sub

                    Dim chosenPath As String = Nothing
                    If displayToPath.TryGetValue(selfile, chosenPath) AndAlso Not String.IsNullOrWhiteSpace(chosenPath) Then
                        SLib.ShowTextFileEditor(chosenPath, $"{AN} FindClause library file '{chosenPath}':", True)
                    End If

                Catch ex As Exception
                    SLib.ShowCustomMessageBox("Error while listing FindClause library files:" & vbCrLf & ex.Message)
                End Try
                Return
            End If

            ' 3) Build segment display list (Title + filename + (local) if local)
            allLibs.Sort(Function(a, b) String.Compare(a.Title, b.Title, StringComparison.OrdinalIgnoreCase))
            Dim segmentDisplayMap As New Dictionary(Of String, ClauseLibrary)(StringComparer.OrdinalIgnoreCase)
            For Each cl In allLibs
                Dim disp As String = $"{cl.Title} [{IO.Path.GetFileName(cl.SourcePath)}]"
                If cl.IsLocal Then disp &= " (local)"
                disp = MakeUniqueDisplay(disp, segmentDisplayMap.Keys)
                segmentDisplayMap(disp) = cl
            Next
            Dim segmentOptions = segmentDisplayMap.Keys.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase).ToList()
            If segmentOptions.Count = 0 Then
                ShowCustomMessageBox("No clause library segments available.")
                Return
            End If

            ' 4) User picks segment + Clean option
            Dim pSeg As New SLib.InputParameter("Clause Segment", segmentOptions(0)) With {.Options = New List(Of String)(segmentOptions)}
            Dim pClean As New SLib.InputParameter("Clean/anonymize", False)
            Dim paramArr() As SLib.InputParameter = {pSeg, pClean}
            If ShowCustomVariableInputForm("Select the library segment to which you want to append the clause:", AN & " AddClause", paramArr) = False Then
                Return
            End If
            Dim chosenDisp As String = CStr(paramArr(0).Value)
            Dim doClean As Boolean = CBool(paramArr(1).Value)

            If String.IsNullOrWhiteSpace(chosenDisp) OrElse Not segmentDisplayMap.ContainsKey(chosenDisp) Then
                ShowCustomMessageBox("Invalid segment selection.")
                Return
            End If
            Dim chosenSegment = segmentDisplayMap(chosenDisp)

            ' 5) Optional clean / anonymize
            Dim finalText As String = selectedText
            If doClean Then
                Try
                    Dim cleaned As String = Await LLM(SP_FindClause_Clean, "<TEXTTOPROCESS>" & selectedText & "</TEXTTOPROCESS>", "", "", 0, False)

                    If String.IsNullOrWhiteSpace(cleaned) Then
                        ShowCustomMessageBox("No cleaned text returned - aborting.")
                        Return
                    End If
                    Dim edited = ShowCustomWindow("The cleaning and anonymization resulted in the following text for your review:", cleaned.Trim(), "Edit text to insert, then press OK or Cancel.", $"{AN} AddClause", True)
                    If String.IsNullOrWhiteSpace(edited) Then
                        ShowCustomMessageBox("Operation cancelled (no text).")
                        Return
                    End If
                    finalText = edited.Trim()
                Catch ex As Exception
                    ShowCustomMessageBox("Clean/anonymize step failed: " & ex.Message)
                    Return
                End Try
            End If
            If String.IsNullOrWhiteSpace(finalText) Then
                ShowCustomMessageBox("No clause text to add (empty).")
                Return
            End If

            ' 6) Update ONLY that segment inside its source file.
            Dim targetFile As String = chosenSegment.SourcePath
            If Not IO.File.Exists(targetFile) Then
                ShowCustomMessageBox("Underlying library file no longer exists.")
                Return
            End If

            Dim success As Boolean = False
            Dim attempt As Integer = 0

            While Not success
                attempt += 1
                Try
                    ' Try to obtain exclusive lock
                    Using fs As New IO.FileStream(targetFile,
                                              IO.FileMode.Open,
                                              IO.FileAccess.ReadWrite,
                                              IO.FileShare.None)

                        ' Read file content through the same locked stream
                        Dim rawText As String
                        Using sr As New IO.StreamReader(fs, New System.Text.UTF8Encoding(False), True, 4096, True)
                            rawText = sr.ReadToEnd()
                        End Using

                        ' Normalize + split into lines (preserve empty file gracefully)
                        Dim allLines As List(Of String)
                        If String.IsNullOrEmpty(rawText) Then
                            allLines = New List(Of String)
                        Else
                            ' Normalize line endings to LF then split; we will write back with CRLF
                            Dim norm = rawText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
                            allLines = norm.Split(New String() {vbLf}, StringSplitOptions.None).ToList()
                        End If

                        ' Parse segments with line positions so we can replace only chosen segment content
                        Dim segments = ParseSegmentsWithPositions(allLines)
                        Dim segInfo = segments.FirstOrDefault(Function(s) s.Title.Equals(chosenSegment.Title, StringComparison.OrdinalIgnoreCase))
                        If segInfo Is Nothing Then
                            ShowCustomMessageBox("Selected segment not found in file (it may have been renamed).")
                            Exit While
                        End If

                        ' Extract original JSON text for this segment (content lines only)
                        Dim originalSegmentJson As String = ""
                        If segInfo.ContentLineCount > 0 Then
                            originalSegmentJson = String.Join(vbCrLf, allLines.Skip(segInfo.ContentStartLine).Take(segInfo.ContentLineCount))
                        End If

                        ' Build new JSON (append record) preserving style
                        Dim updatedSegmentJson As String = AppendClauseToSegment(originalSegmentJson, finalText)
                        If String.IsNullOrWhiteSpace(updatedSegmentJson) Then
                            ShowCustomMessageBox("Failed to build updated JSON for the selected segment.")
                            Exit While
                        End If

                        ' Remove old content lines
                        For i = 1 To segInfo.ContentLineCount
                            allLines.RemoveAt(segInfo.ContentStartLine)
                        Next
                        ' Insert new content lines
                        Dim newLines = Split(updatedSegmentJson, vbCrLf).ToList()
                        allLines.InsertRange(segInfo.ContentStartLine, newLines)

                        ' Rewind & overwrite file
                        fs.Position = 0
                        fs.SetLength(0)
                        Using sw As New IO.StreamWriter(fs, New System.Text.UTF8Encoding(False), 4096, True)
                            For i = 0 To allLines.Count - 1
                                If i > 0 Then sw.Write(vbCrLf)
                                sw.Write(allLines(i))
                            Next
                            sw.Flush()
                        End Using
                    End Using

                    success = True
                    ShowCustomMessageBox($"Clause added to segment '{chosenSegment.Title}' of your library file.",
                                                                extraButtonText:="Edit file",
                                                                extraButtonAction:=Sub()
                                                                                       SLib.ShowTextFileEditor(targetFile, $"Edit your library file ('{targetFile}'):", True, _context)
                                                                                   End Sub)

                Catch ioEx As IO.IOException
                    ' Sharing violation only happens on opening, not during internal re-read.
                    Dim choice = ShowCustomYesNoBox($"Could not acquire exclusive access (attempt {attempt}). Retry?", "Retry", "Abort")
                    If choice <> 1 Then
                        ShowCustomMessageBox("Operation aborted - could not acquire lock.")
                        Return
                    End If
                    System.Threading.Thread.Sleep(250)

                Catch ex As Exception
                    ShowCustomMessageBox("Error updating segment: " & ex.Message)
                    Return
                End Try
            End While

        Catch ex As System.Exception
#If DEBUG Then
        System.Diagnostics.Debug.WriteLine("AddClause error: " & ex.Message)
        System.Diagnostics.Debug.WriteLine(ex.StackTrace)
#End If
            ShowCustomMessageBox("Error in AddClause: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Parses the entire library file content (provided as a list of lines) into segments with precise line positions.
    ''' </summary>
    ''' <param name="lines">List of text lines from the library file.</param>
    ''' <returns>List of SegmentInfo objects describing each segment's title and content range.</returns>
    ''' <remarks>
    ''' This function ignores SP_FindClause and SP_MergePrompt override lines when calculating content boundaries.
    ''' Segments are delimited by lines matching the pattern [SegmentTitle].
    ''' </remarks>
    Private Function ParseSegmentsWithPositions(lines As List(Of String)) As List(Of SegmentInfo)
        Dim result As New List(Of SegmentInfo)()
        Dim currentTitle As String = Nothing
        Dim contentStart As Integer = -1

        ' Ignore SP_FindClause / SP_MergePrompt lines from content (same logic as original parser)
        Dim i As Integer = 0
        While i < lines.Count
            Dim rawLine = lines(i)
            Dim line = (If(rawLine, "")).Trim()

            Dim isHeader As Boolean = line.StartsWith("[") AndAlso line.EndsWith("]")
            If isHeader Then
                ' Flush previous segment
                If currentTitle IsNot Nothing Then
                    Dim segEndLineExclusive = i ' current header line starts next segment
                    Dim contentCount = System.Math.Max(0, segEndLineExclusive - contentStart)
                    result.Add(New SegmentInfo With {
                    .Title = currentTitle,
                    .ContentStartLine = contentStart,
                    .ContentLineCount = contentCount
                })
                End If

                currentTitle = line.Substring(1, line.Length - 2).Trim()
                contentStart = i + 1 ' content starts after header
                i += 1
                Continue While
            End If

            ' Skip prompt override lines from segment content (they are not part of JSON)
            If currentTitle IsNot Nothing AndAlso
           (Regex.IsMatch(line, "^\s*SP_FindClause\s*=", RegexOptions.IgnoreCase) OrElse
            Regex.IsMatch(line, "^\s*SP_MergePrompt\s*=", RegexOptions.IgnoreCase)) Then
                ' Ensure contentStart moves forward if prompt lines appear at the beginning
                If contentStart = i Then contentStart = i + 1
            End If

            i += 1
        End While

        ' Flush last segment
        If currentTitle IsNot Nothing Then
            Dim segEnd = lines.Count
            Dim contentCount = System.Math.Max(0, segEnd - contentStart)
            result.Add(New SegmentInfo With {
            .Title = currentTitle,
            .ContentStartLine = contentStart,
            .ContentLineCount = contentCount
        })
        End If

        Return result
    End Function

    ''' <summary>
    ''' Contains metadata about a clause library segment's position within the library file.
    ''' </summary>
    Private Class SegmentInfo
        ''' <summary>Gets or sets the segment title.</summary>
        Public Property Title As String
        ''' <summary>Gets or sets the zero-based line index where JSON content begins.</summary>
        Public Property ContentStartLine As Integer
        ''' <summary>Gets or sets the number of lines belonging to JSON content.</summary>
        Public Property ContentLineCount As Integer
    End Class

    ''' <summary>
    ''' Appends the provided clause text to the existing segment JSON, preserving the original format.
    ''' </summary>
    ''' <param name="originalJson">The JSON content of the target segment.</param>
    ''' <param name="finalText">The clause text to append.</param>
    ''' <returns>Updated JSON string with the new clause appended, or Nothing if the operation fails or is cancelled.</returns>
    ''' <remarks>
    ''' Supports three storage formats:
    ''' (A) Wrapper object with a "Records" array.
    ''' (B) Pure JSON array.
    ''' (C) Sequence of standalone objects separated by blank lines.
    ''' Performs duplicate checking based on the dynamically detected field name.
    ''' </remarks>
    Private Function AppendClauseToSegment(originalJson As String, finalText As String) As String
        Dim trimmed = (If(originalJson, "")).Trim()

        ' Case A: Wrapper object containing "Records" array
        If LooksLikeWrapperWithRecords(trimmed) Then
            Try
                Dim obj = JObject.Parse(trimmed)
                Dim arr = TryCast(obj("Records"), JArray)
                If arr Is Nothing Then
                    arr = New JArray()
                    obj("Records") = arr
                End If

                Dim fieldName = DetectFieldNameFromLast(arr)
                If String.IsNullOrWhiteSpace(fieldName) Then fieldName = "Text"

                ' Duplicate check
                Dim dup = arr.OfType(Of JObject)().Any(Function(o) HasStringValue(o, fieldName, finalText))
                If dup Then
                    Dim c = ShowCustomYesNoBox("A record with identical text already exists in this segment. Add anyway?", "Add duplicate", "Abort")
                    If c <> 1 Then Return Nothing
                End If

                Dim newRec As New JObject From {{fieldName, finalText}}
                arr.Add(newRec)
                Return obj.ToString(Formatting.Indented)
            Catch
                ' Fall through to next styles
            End Try
        End If

        ' Case B: Pure JSON array
        If trimmed.StartsWith("[") Then
            Try
                Dim arr = JArray.Parse(trimmed)
                Dim fieldName = DetectFieldNameFromLast(arr)
                If String.IsNullOrWhiteSpace(fieldName) Then fieldName = "Text"

                Dim dup = arr.OfType(Of JObject)().Any(Function(o) HasStringValue(o, fieldName, finalText))
                If dup Then
                    Dim c = ShowCustomYesNoBox("A record with identical text already exists in this segment. Add anyway?", "Add duplicate", "Abort")
                    If c <> 1 Then Return Nothing
                End If

                arr.Add(New JObject From {{fieldName, finalText}})
                Return arr.ToString(Formatting.Indented)
            Catch
                ' Fall through
            End Try
        End If

        ' Case C: Sequence of standalone objects
        Dim objects = ParseStandaloneObjectSequence(trimmed)
        If objects IsNot Nothing Then
            Dim fieldName = DetectFieldNameFromLast(objects)
            If String.IsNullOrWhiteSpace(fieldName) Then fieldName = "Text"

            Dim dup = objects.OfType(Of JObject)().Any(Function(o) HasStringValue(o, fieldName, finalText))
            If dup Then
                Dim c = ShowCustomYesNoBox("A record with identical text already exists in this segment. Add anyway?", "Add duplicate", "Abort")
                If c <> 1 Then Return Nothing
            End If

            objects.Add(New JObject From {{fieldName, finalText}})

            ' Reconstruct sequence (keep style: each object pretty printed separated by blank line)
            Dim sb As New System.Text.StringBuilder()
            For i = 0 To objects.Count - 1
                If i > 0 Then sb.AppendLine().AppendLine()
                sb.Append(objects(i).ToString(Formatting.Indented))
            Next
            Return sb.ToString()
        End If

        ' Empty segment: create new wrapper
        If String.IsNullOrWhiteSpace(trimmed) Then
            Dim arr As New JArray()
            arr.Add(New JObject From {{"Text", finalText}})
            Return "{""Records"":" & arr.ToString(Formatting.None) & "}"
        End If

        ' Fallback: treat entire content as single object and convert to Records array
        Try
            Dim singleObj = JObject.Parse(trimmed)
            Dim fieldName As String = singleObj.Properties().FirstOrDefault(Function(p) p.Value.Type = JTokenType.String)?.Name
            If String.IsNullOrWhiteSpace(fieldName) Then fieldName = "Text"
            Dim arr As New JArray(singleObj, New JObject From {{fieldName, finalText}})
            Return "{""Records"":" & arr.ToString(Formatting.None) & "}"
        Catch
            ' Unable to process
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Determines whether the provided string appears to be a JSON object containing a "Records" array.
    ''' </summary>
    ''' <param name="s">JSON string to inspect.</param>
    ''' <returns>True if the string starts with "{" and contains a "Records" property.</returns>
    Private Function LooksLikeWrapperWithRecords(s As String) As Boolean
        Return s.StartsWith("{") AndAlso s.IndexOf("""Records""", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    ''' <summary>
    ''' Detects the clause field name by examining the last object in the provided container.
    ''' </summary>
    ''' <param name="container">IEnumerable of JObjects (JArray or List(Of JObject)).</param>
    ''' <returns>The name of the first string property in the last object, or Nothing if not found.</returns>
    ''' <remarks>
    ''' This method is used to determine the dynamic field name for clause text storage.
    ''' </remarks>
    Private Function DetectFieldNameFromLast(container As IEnumerable) As String
        Dim lastObj As JObject = Nothing
        For Each o In container
            If TypeOf o Is JObject Then lastObj = CType(o, JObject)
        Next
        If lastObj Is Nothing Then Return Nothing
        For Each p In lastObj.Properties()
            If p.Value.Type = JTokenType.String Then Return p.Name
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Checks whether the specified JObject has a string property with the given field name and value.
    ''' </summary>
    ''' <param name="o">JObject to inspect.</param>
    ''' <param name="fieldName">Property name to check.</param>
    ''' <param name="value">Expected string value.</param>
    ''' <returns>True if the object contains the field with an exact string match (ordinal comparison).</returns>
    Private Function HasStringValue(o As JObject, fieldName As String, value As String) As Boolean
        Dim tok = o(fieldName)
        Return tok IsNot Nothing AndAlso tok.Type = JTokenType.String AndAlso String.Equals(CStr(tok), value, StringComparison.Ordinal)
    End Function

    ''' <summary>
    ''' Attempts to parse a sequence of standalone JSON objects separated by blank lines.
    ''' </summary>
    ''' <param name="raw">Raw JSON string potentially containing multiple objects.</param>
    ''' <returns>List of parsed JObjects, or Nothing if the content is not a valid standalone sequence.</returns>
    ''' <remarks>
    ''' A valid sequence contains multiple objects separated by one or more blank lines. Single-object
    ''' content is treated as a fallback case and returns Nothing.
    ''' </remarks>
    Private Function ParseStandaloneObjectSequence(raw As String) As List(Of JObject)
        If String.IsNullOrWhiteSpace(raw) Then Return New List(Of JObject)()
        Dim parts As New List(Of String)()

        ' Split on blank-line boundaries that separate top-level objects
        Dim regexSplit = New Regex("(?<=\})(?:\s*\r?\n){1,}(?=\{)", RegexOptions.Singleline)
        Dim rawParts = regexSplit.Split(raw)

        ' If only one part, ensure it is NOT just a single object (then it's not a sequence style)
        If rawParts.Length = 1 Then
            ' If it parses as object, we treat that in fallback, so return Nothing here
            Try
                JObject.Parse(rawParts(0))
                Return Nothing
            Catch
                ' Not valid JSON => maybe malformed -> ignore
                Return Nothing
            End Try
        End If

        Dim list As New List(Of JObject)
        For Each p In rawParts
            Dim t = p.Trim()
            If t = "" Then Continue For
            Try
                list.Add(JObject.Parse(t))
            Catch
                ' If any object fails to parse -> not reliable -> abort
                Return Nothing
            End Try
        Next
        Return list
    End Function

End Class
