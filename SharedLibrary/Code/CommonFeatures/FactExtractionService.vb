' Part of: Red Ink Shared Library
' Copyright by David Rosenthal
Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    Public Module FactExtractionService

        <DebuggerDisplay("{Name} ({Type})")>
        Public Class ExtractionSchemaColumn
            Public Property Name As String
            Public Property Type As String ' text | date | datetime | time | number | other
        End Class

        Public Class ExtractionRow
            Public Property Values As System.Collections.Generic.List(Of Object)
        End Class

        Public Class FactExtractionAggregateResult
            Public Property Schema As System.Collections.Generic.List(Of ExtractionSchemaColumn)
            Public Property Rows As System.Collections.Generic.List(Of ExtractionRow)
            Public Property Errors As System.Collections.Generic.List(Of String)
            Public Property ProcessedFiles As Integer
            Public Property FailedFiles As Integer
            Public Property FailedFileNames As System.Collections.Generic.List(Of String)
            Public Property SourceDirectory As String
        End Class

        Public Const Setting_ManualInstruction As String = "Extraction_ManualInstruction"
        Public Const Setting_DateColumns As String = "Extraction_DateColumns"
        Public Const Setting_SortColumn As String = "Extraction_SortColumn"
        Public Const Setting_SortDirection As String = "Extraction_SortDirection"
        Public Const Setting_DoOcr As String = "Extraction_DoOcr"
        Public Const Setting_DateClampFrom As String = "Extraction_DateClampFrom"
        Public Const Setting_DateClampTo As String = "Extraction_DateClampTo"
        Public Const Setting_OutputLanguage As String = "Extraction_OutputLanguage"
        Public Const Setting_DateOutputFormat As String = "Extraction_DateOutputFormat"


        Public Function ParseFlexibleDate(raw As String) As Date?
            If String.IsNullOrWhiteSpace(raw) Then Return Nothing
            Dim t = raw.Trim()

            Dim mShort = System.Text.RegularExpressions.Regex.Match(t, "^(19|20)\d{2}-([1-9]|1[0-2])$")
            If mShort.Success Then
                Dim yearPart = mShort.Value.Substring(0, 4)
                Dim monthPart = mShort.Groups(2).Value.PadLeft(2, "0"c)
                Return New Date(CInt(yearPart), CInt(monthPart), 1)
            End If

            Dim isoMonth = System.Text.RegularExpressions.Regex.Match(t, "^(19|20)\d{2}-(0[1-9]|1[0-2])$")
            If isoMonth.Success Then
                Dim yearPart = t.Substring(0, 4)
                Dim monthPart = t.Substring(5, 2)
                Return New Date(CInt(yearPart), CInt(monthPart), 1)
            End If

            Dim isoFull = System.Text.RegularExpressions.Regex.Match(t, "^(19|20)\d{2}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$")
            If isoFull.Success Then
                Try : Return Date.ParseExact(t, "yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture) : Catch : End Try
            End If

            Dim dmY = System.Text.RegularExpressions.Regex.Match(t, "^(?<d>[0-3]?\d)\.(?<m>[0-1]?\d)\.(?<y>(19|20)\d{2})$")
            If dmY.Success Then
                Try
                    Return New Date(CInt(dmY.Groups("y").Value), CInt(dmY.Groups("m").Value), CInt(dmY.Groups("d").Value))
                Catch
                End Try
            End If

            Dim dmYY = System.Text.RegularExpressions.Regex.Match(t, "^(?<d>[0-3]?\d)\.(?<m>[0-1]?\d)\.(?<y>\d{2})$")
            If dmYY.Success Then
                Try
                    Dim yy = CInt(dmYY.Groups("y").Value)
                    Dim y = If(yy < 50, 2000 + yy, 1900 + yy)
                    Return New Date(y, CInt(dmYY.Groups("m").Value), CInt(dmYY.Groups("d").Value))
                Catch
                End Try
            End If

            Dim monthYear = System.Text.RegularExpressions.Regex.Match(t, "^(?<m>\p{L}+)\s+(?<y>(19|20)\d{2})$")
            If monthYear.Success Then
                Try
                    Dim y = CInt(monthYear.Groups("y").Value)
                    Dim monthName = monthYear.Groups("m").Value
                    Dim m = DateTime.ParseExact(monthName, "MMMM", Globalization.CultureInfo.InvariantCulture).Month
                    Return New Date(y, m, 1)
                Catch
                End Try
            End If

            Dim yearOnly = System.Text.RegularExpressions.Regex.Match(t, "^(19|20)\d{2}$")
            If yearOnly.Success Then Return New Date(CInt(yearOnly.Value), 1, 1)

            Dim dt As DateTime
            If DateTime.TryParse(t, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.AllowWhiteSpaces, dt) Then Return dt
            If DateTime.TryParse(t, Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.AllowWhiteSpaces, dt) Then Return dt
            Return Nothing
        End Function

        Public Function NormalizeDate(raw As String) As String
            Dim p = ParseFlexibleDate(raw)
            If Not p.HasValue Then Return raw
            Dim t = raw.Trim()
            If System.Text.RegularExpressions.Regex.IsMatch(t, "^(19|20)\d{2}$") Then Return p.Value.ToString("yyyy", Globalization.CultureInfo.InvariantCulture)
            If System.Text.RegularExpressions.Regex.IsMatch(t, "^(19|20)\d{2}-(0?[1-9]|1[0-2])$") _
               OrElse System.Text.RegularExpressions.Regex.IsMatch(t, "^\p{L}+\s+(19|20)\d{2}$") Then
                Return p.Value.ToString("yyyy-MM", Globalization.CultureInfo.InvariantCulture)
            End If
            If System.Text.RegularExpressions.Regex.IsMatch(t, "^[0-3]?\d\.[0-1]?\d\.(\d{2}|\d{4})$") Then Return p.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
            If System.Text.RegularExpressions.Regex.IsMatch(t, "^(19|20)\d{2}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$") Then Return p.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
            Return p.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Function ToComparableDate(normalized As String) As Date?
            If String.IsNullOrWhiteSpace(normalized) Then Return Nothing
            Dim s = normalized.Trim()
            Dim yearOnly = System.Text.RegularExpressions.Regex.Match(s, "^(?<year>(19|20)\d{2})$")
            If yearOnly.Success Then Return New Date(CInt(yearOnly.Groups("year").Value), 1, 1)
            Dim yearMonth = System.Text.RegularExpressions.Regex.Match(s, "^(?<year>(19|20)\d{2})-(?<month>(0[1-9]|1[0-2]))$")
            If yearMonth.Success Then Return New Date(CInt(yearMonth.Groups("year").Value), CInt(yearMonth.Groups("month").Value), 1)
            Return ParseFlexibleDate(s)
        End Function

        Public Function ParseSingleFileJson(json As String) As (schema As System.Collections.Generic.List(Of ExtractionSchemaColumn), rows As System.Collections.Generic.List(Of ExtractionRow), fileName As String, notes As String)
            Dim schema As New System.Collections.Generic.List(Of ExtractionSchemaColumn)
            Dim rows As New System.Collections.Generic.List(Of ExtractionRow)
            Dim fileName As String = ""
            Dim notes As String = ""
            If String.IsNullOrWhiteSpace(json) Then Return (schema, rows, fileName, notes)
            Try
                Dim jt = JToken.Parse(json)
                Dim schemaTok = jt("schema")
                If schemaTok IsNot Nothing AndAlso schemaTok.Type = JTokenType.Array Then
                    For Each c In schemaTok
                        Dim name = CStr(c("name"))
                        Dim typ = CStr(c("type"))
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        schema.Add(New ExtractionSchemaColumn With {.Name = name.Trim(), .Type = If(String.IsNullOrWhiteSpace(typ), "text", typ.Trim().ToLowerInvariant())})
                    Next
                End If
                Dim rowsTok = jt("rows")
                If rowsTok IsNot Nothing AndAlso rowsTok.Type = JTokenType.Array Then
                    For Each r In rowsTok
                        Dim er As New ExtractionRow With {.Values = New System.Collections.Generic.List(Of Object)()}
                        If r.Type = JTokenType.Array Then
                            For Each v In DirectCast(r, JArray)
                                er.Values.Add(ConvertToken(v))
                            Next
                        ElseIf r.Type = JTokenType.Object Then
                            For Each col In schema
                                er.Values.Add(ConvertToken(r(col.Name)))
                            Next
                        End If
                        If er.Values.Count > 0 Then rows.Add(er)
                    Next
                End If
                fileName = CStr(jt("file_name"))
                notes = CStr(jt("notes"))
            Catch
            End Try
            Return (schema, rows, fileName, notes)
        End Function

        Private Function ConvertToken(tok As JToken) As Object
            If tok Is Nothing Then Return ""
            Select Case tok.Type
                Case JTokenType.String : Return CStr(tok)
                Case JTokenType.Integer : Return CInt(tok)
                Case JTokenType.Float : Return CDbl(tok)
                Case JTokenType.Boolean : Return CBool(tok)
                Case Else : Return tok.ToString()
            End Select
        End Function

        Public Function ParseUserSchemaSpec(spec As String) As System.Collections.Generic.List(Of ExtractionSchemaColumn)
            Dim result As New System.Collections.Generic.List(Of ExtractionSchemaColumn)
            If String.IsNullOrWhiteSpace(spec) Then Return result
            For Each raw In spec.Split(";"c)
                Dim token = raw.Trim()
                If token.Length = 0 Then Continue For
                Dim namePart = token
                Dim typePart = "text"
                Dim colonIdx = token.IndexOf(":"c)
                If colonIdx >= 0 Then
                    namePart = token.Substring(0, colonIdx).Trim()
                    Dim after = token.Substring(colonIdx + 1).Trim()
                    If after.EndsWith("*") Then after = after.Substring(0, after.Length - 1).Trim()
                    If after.Length > 0 Then typePart = after.ToLowerInvariant()
                ElseIf token.EndsWith("*") Then
                    namePart = token.Substring(0, token.Length - 1).Trim()
                End If
                If String.IsNullOrWhiteSpace(namePart) Then Continue For
                result.Add(New ExtractionSchemaColumn With {.Name = namePart, .Type = typePart})
            Next
            Return result
        End Function

        Public Function DetectSortColumnFromSpec(spec As String) As Integer
            If String.IsNullOrWhiteSpace(spec) Then Return 0
            Dim idx = 0
            For Each raw In spec.Split(";"c)
                Dim token = raw.Trim()
                If token.Length = 0 Then Continue For
                idx += 1
                If token.EndsWith("*") Then Return idx
                Dim colonIdx = token.IndexOf(":"c)
                If colonIdx >= 0 Then
                    Dim after = token.Substring(colonIdx + 1).Trim()
                    If after.EndsWith("*") Then Return idx
                End If
            Next
            Return 0
        End Function

        Public Async Function GenerateSchemaFromAiAsync(instruction As String,
                                                        interpolateSystemPromptFunc As Func(Of String, String),
                                                        llmFunc As Func(Of String, String, String, String, Integer, Boolean, Boolean, Threading.Tasks.Task(Of String)),
                                                        useSecondApi As Boolean,
                                                        context As ISharedContext) As Threading.Tasks.Task(Of System.Collections.Generic.List(Of ExtractionSchemaColumn))
            Dim userText = ""
            Dim systemPrompt = interpolateSystemPromptFunc(context.SP_ExtractSchema)
            Dim jsonResp = Await llmFunc(systemPrompt, userText, "", "", 0, useSecondApi, False)
            jsonResp = WebAgentInterpreter.SanitizeLlmResult(jsonResp)
            Dim schemaOnly As New System.Collections.Generic.List(Of ExtractionSchemaColumn)
            If String.IsNullOrWhiteSpace(jsonResp) Then Return schemaOnly
            Try
                Dim jt = JToken.Parse(jsonResp)
                Dim st = jt("schema")
                If st IsNot Nothing AndAlso st.Type = JTokenType.Array Then
                    For Each c In st
                        Dim name = CStr(c("name"))
                        Dim typ = CStr(c("type"))
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        schemaOnly.Add(New ExtractionSchemaColumn With {
                            .Name = name.Trim(),
                            .Type = If(String.IsNullOrWhiteSpace(typ), "text", typ.Trim().ToLowerInvariant())
                        })
                    Next
                End If
            Catch
            End Try
            Return schemaOnly
        End Function

        Public Function BuildConstrainedSystemPrompt(originalInterpolatedPrompt As String,
                                                     fixedSchema As System.Collections.Generic.List(Of ExtractionSchemaColumn)) As String
            If fixedSchema Is Nothing OrElse fixedSchema.Count = 0 Then Return originalInterpolatedPrompt
            Dim sb As New StringBuilder(originalInterpolatedPrompt.Length + 300)
            sb.AppendLine(originalInterpolatedPrompt)
            sb.AppendLine()
            sb.AppendLine("FIXED ORDERED SCHEMA (DO NOT RENAME OR REORDER):")
            sb.AppendLine(String.Join(" | ", fixedSchema.Select(Function(c) c.Name)))
            sb.AppendLine("Return JSON: {""rows"":[[...]],""file_name"":""<file>""} ONLY. No schema array, no commentary.")
            Return sb.ToString()
        End Function

        Public Sub MergeIntoAggregate(master As FactExtractionAggregateResult,
                                      schema As System.Collections.Generic.List(Of ExtractionSchemaColumn),
                                      rows As System.Collections.Generic.List(Of ExtractionRow),
                                      sourceFileName As String,
                                      dateColumnsUser As System.Collections.Generic.List(Of Integer))

            If master.Schema.Count = 0 AndAlso schema.Count > 0 Then
                If Not schema.Any(Function(c) c.Name.Equals("File", StringComparison.OrdinalIgnoreCase)) Then
                    schema.Add(New ExtractionSchemaColumn With {.Name = "File", .Type = "text"})
                End If
                master.Schema.AddRange(schema)
            Else
                For Each c In schema
                    If Not master.Schema.Any(Function(mc) mc.Name.Equals(c.Name, StringComparison.OrdinalIgnoreCase)) Then
                        master.Schema.Add(New ExtractionSchemaColumn With {.Name = c.Name, .Type = c.Type})
                    End If
                Next
                If Not master.Schema.Any(Function(mc) mc.Name.Equals("File", StringComparison.OrdinalIgnoreCase)) Then
                    master.Schema.Add(New ExtractionSchemaColumn With {.Name = "File", .Type = "text"})
                End If
            End If

            Dim fileColIndex = master.Schema.FindIndex(Function(c) c.Name.Equals("File", StringComparison.OrdinalIgnoreCase))
            If fileColIndex < 0 Then
                master.Schema.Add(New ExtractionSchemaColumn With {.Name = "File", .Type = "text"})
                fileColIndex = master.Schema.Count - 1
            End If

            For Each r In rows
                Dim newRow As New ExtractionRow With {.Values = New System.Collections.Generic.List(Of Object)()}
                For i = 0 To master.Schema.Count - 1
                    Dim v As Object = ""
                    If i < r.Values.Count Then v = r.Values(i)
                    newRow.Values.Add(v)
                Next
                newRow.Values(fileColIndex) = sourceFileName
                For Each dc In dateColumnsUser
                    Dim idx = dc - 1
                    If idx >= 0 AndAlso idx < newRow.Values.Count Then
                        Dim raw = CStr(newRow.Values(idx))
                        If Not String.IsNullOrWhiteSpace(raw) Then
                            newRow.Values(idx) = NormalizeDate(raw)
                        End If
                    End If
                Next
                master.Rows.Add(newRow)
            Next
        End Sub

        Private Sub ApplyDateClamps(result As FactExtractionAggregateResult,
                                    dateColumnsUser As System.Collections.Generic.List(Of Integer),
                                    clampFromRaw As String,
                                    clampToRaw As String)
            If result Is Nothing OrElse result.Rows.Count = 0 OrElse dateColumnsUser Is Nothing OrElse dateColumnsUser.Count = 0 Then Return
            Dim clampFrom = ParseFlexibleDate(clampFromRaw)
            Dim clampTo = ParseFlexibleDate(clampToRaw)
            If Not clampFrom.HasValue AndAlso Not clampTo.HasValue Then Return

            Dim keep As New System.Collections.Generic.List(Of ExtractionRow)
            For Each row In result.Rows
                Dim ok As Boolean = True
                For Each dc In dateColumnsUser
                    Dim idx = dc - 1
                    If idx < 0 OrElse idx >= row.Values.Count Then Continue For
                    Dim cell = CStr(row.Values(idx))
                    Dim dt = ToComparableDate(NormalizeDate(cell))
                    If dt.HasValue Then
                        If clampFrom.HasValue AndAlso dt.Value < clampFrom.Value Then ok = False : Exit For
                        If clampTo.HasValue AndAlso dt.Value > clampTo.Value Then ok = False : Exit For
                    End If
                Next
                If ok Then keep.Add(row)
            Next
            result.Rows = keep
        End Sub

        Public Sub SortAggregate(result As FactExtractionAggregateResult,
                                 sortColumn As Integer,
                                 sortDir As String)
            If result Is Nothing OrElse result.Rows.Count = 0 Then Return
            If sortColumn <= 0 Then Return
            Dim idx = sortColumn - 1
            If idx >= result.Schema.Count Then Return
            Dim typeHint = result.Schema(idx).Type
            Dim asc = Not sortDir.Equals("DESC", StringComparison.OrdinalIgnoreCase)
            result.Rows.Sort(Function(a, b)
                                 Dim av = If(idx < a.Values.Count, a.Values(idx), Nothing)
                                 Dim bv = If(idx < b.Values.Count, b.Values(idx), Nothing)
                                 Dim cmp = CompareValues(av, bv, typeHint)
                                 Return If(asc, cmp, -cmp)
                             End Function)
        End Sub

        Private Function CompareValues(a As Object, b As Object, typeHint As String) As Integer
            If a Is Nothing AndAlso b Is Nothing Then Return 0
            If a Is Nothing Then Return -1
            If b Is Nothing Then Return 1
            Dim sa = a.ToString()
            Dim sb = b.ToString()

            If {"date", "datetime", "time"}.Contains(typeHint) Then
                Dim da = ToComparableDate(NormalizeDate(sa))
                Dim db = ToComparableDate(NormalizeDate(sb))
                If da.HasValue AndAlso db.HasValue Then Return DateTime.Compare(da.Value, db.Value)
            End If

            Dim daNum As Double
            Dim dbNum As Double
            If Double.TryParse(sa, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, daNum) AndAlso
               Double.TryParse(sb, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, dbNum) Then
                Return daNum.CompareTo(dbNum)
            End If

            Return String.Compare(sa, sb, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Async Function RunFactExtractionAsync(filePaths As System.Collections.Generic.List(Of String),
                                             instruction As String,
                                             dateColumnsUser As System.Collections.Generic.List(Of Integer),
                                             sortColumn As Integer,
                                             sortDirection As String,
                                             doOcr As Boolean,
                                             useSecondApi As Boolean,
                                             sourceDirectory As String,
                                             interpolateSystemPromptFunc As Func(Of String, String),
                                             llmFunc As Func(Of String, String, String, String, Integer, Boolean, Boolean, Threading.Tasks.Task(Of String)),
                                             GetFileContentFunc As Func(Of String, Boolean, Boolean, Boolean, Threading.Tasks.Task(Of String)),
                                             context As ISharedContext,
                                             Optional fixedSchema As System.Collections.Generic.List(Of ExtractionSchemaColumn) = Nothing,
                                             Optional clampFrom As String = Nothing,
                                             Optional clampTo As String = Nothing,
                                             Optional progressCallback As Action(Of Integer, Integer, String) = Nothing,
                                             Optional mergeDateColumn As Integer = 0,
                                             Optional mergeRowsViaLlm As Boolean = False,
                                             Optional mergeInstruction As String = Nothing) _
                                             As Threading.Tasks.Task(Of FactExtractionAggregateResult)

            Dim agg As New FactExtractionAggregateResult With {
                .Schema = New System.Collections.Generic.List(Of ExtractionSchemaColumn),
                .Rows = New System.Collections.Generic.List(Of ExtractionRow),
                .Errors = New System.Collections.Generic.List(Of String),
                .FailedFileNames = New System.Collections.Generic.List(Of String),
                .SourceDirectory = sourceDirectory
            }

            If filePaths Is Nothing OrElse filePaths.Count = 0 Then
                agg.Errors.Add("No input files.")
                Return agg
            End If

            For i = 0 To filePaths.Count - 1
                Dim path = filePaths(i)
                If progressCallback IsNot Nothing Then
                    progressCallback(i, filePaths.Count, "Processing " & System.IO.Path.GetFileName(path) & " (" & (i + 1).ToString() & " of " & filePaths.Count.ToString() & ")")
                End If

                If Not File.Exists(path) Then
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End If
                Dim text = Await GetFileContentFunc(path, False, doOcr, False)
                If String.IsNullOrWhiteSpace(text) Then
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End If
                Dim userText = "<TEXTTOPROCESS>" & text & "</TEXTTOPROCESS>"

                Dim sysPrompt = interpolateSystemPromptFunc(context.SP_Extract)
                If fixedSchema IsNot Nothing AndAlso fixedSchema.Count > 0 Then
                    sysPrompt = BuildConstrainedSystemPrompt(sysPrompt, fixedSchema)
                End If
                Dim jsonResp As String = Nothing
                Try
                    jsonResp = Await llmFunc(sysPrompt, userText, "", "", 0, useSecondApi, False)
                    jsonResp = WebAgentInterpreter.SanitizeLlmResult(jsonResp)
                Catch ex As Exception
                    agg.Errors.Add("LLM call failed for '" & System.IO.Path.GetFileName(path) & "': " & ex.Message)
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End Try
                If String.IsNullOrWhiteSpace(jsonResp) Then
                    agg.Errors.Add("Empty AI response for '" & System.IO.Path.GetFileName(path) & "'.")
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End If
                Dim parsed = ParseSingleFileJson(jsonResp)
                If fixedSchema IsNot Nothing AndAlso fixedSchema.Count > 0 AndAlso parsed.schema.Count = 0 Then
                    parsed.schema.AddRange(fixedSchema.Select(Function(c) New ExtractionSchemaColumn With {.Name = c.Name, .Type = c.Type}))
                End If
                If parsed.schema.Count = 0 OrElse parsed.rows.Count = 0 Then
                    agg.Errors.Add("No rows/schema parsed for '" & System.IO.Path.GetFileName(path) & "'.")
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End If
                MergeIntoAggregate(agg, parsed.schema, parsed.rows, System.IO.Path.GetFileName(path), dateColumnsUser)
                agg.ProcessedFiles += 1
            Next

            ApplyDateClamps(agg, dateColumnsUser, clampFrom, clampTo)

            ' Perform merging if requested
            If mergeRowsViaLlm AndAlso mergeDateColumn > 0 Then
                If progressCallback IsNot Nothing Then
                    progressCallback(filePaths.Count, filePaths.Count, "Merging rows by date...")
                End If
                Await MergeRowsByDateAsync(agg,
                               mergeDateColumn,
                               mergeInstruction,
                               useSecondApi,
                               interpolateSystemPromptFunc,
                               llmFunc,
                               context,
                               baseCur:=filePaths.Count,
                               baseTotal:=filePaths.Count,
                               progressCallback:=progressCallback)
            End If

            agg.FailedFiles = agg.FailedFileNames.Count
            If sortColumn > 0 Then SortAggregate(agg, sortColumn, sortDirection)

            If progressCallback IsNot Nothing Then
                progressCallback(filePaths.Count, filePaths.Count, "Completed.")
            End If

            Return agg
        End Function


        ' Update signature of MergeRowsByDateAsync to accept the callback and base values, and emit progress per group
        Private Async Function MergeRowsByDateAsync(agg As FactExtractionAggregateResult,
                                            dateColumn As Integer,
                                            mergeInstruction As String,
                                            useSecondApi As Boolean,
                                            interpolateSystemPromptFunc As Func(Of String, String),
                                            llmFunc As Func(Of String, String, String, String, Integer, Boolean, Boolean, Threading.Tasks.Task(Of String)),
                                            context As ISharedContext,
                                            Optional baseCur As Integer = 0,
                                            Optional baseTotal As Integer = 0,
                                            Optional progressCallback As Action(Of Integer, Integer, String) = Nothing) As Threading.Tasks.Task
            If agg Is Nothing OrElse agg.Rows.Count = 0 Then Return
            If dateColumn <= 0 Then Return
            Dim dateIdx = dateColumn - 1
            If dateIdx >= agg.Schema.Count Then Return

            ' Group rows by normalized date
            Dim groups = agg.Rows.
        GroupBy(Function(r) GetNormalizedDateValue(r, dateIdx)).
        Where(Function(g) Not String.IsNullOrWhiteSpace(g.Key)).
        ToList()

            If groups.Count = 0 Then Return

            Dim newRows As New System.Collections.Generic.List(Of ExtractionRow)
            Dim schemaNames = agg.Schema.Select(Function(c) c.Name).ToList()

            Dim totalGroups = groups.Count
            Dim groupIndex As Integer = 0

            For Each g In groups
                groupIndex += 1
                If progressCallback IsNot Nothing Then
                    Dim label As String = $"Merging rows by date... ({groupIndex}/{totalGroups}) [{g.Key}]"
                    ' Keep bar position at baseCur/baseTotal so it does not jump, but label shows merge progress.
                    progressCallback(baseCur, baseTotal, label)
                End If

                Dim groupRows = g.ToList()
                ' Build JSON input for LLM
                Dim rowsArray As New System.Text.StringBuilder()
                rowsArray.Append("{""date"":""" & g.Key & """,""schema"":[""" & String.Join(""" ,""", schemaNames) & """],""rows"":[")
                For ri = 0 To groupRows.Count - 1
                    Dim r = groupRows(ri)
                    rowsArray.Append("[")
                    For ci = 0 To agg.Schema.Count - 1
                        Dim val As String = ""
                        If ci < r.Values.Count AndAlso r.Values(ci) IsNot Nothing Then
                            val = r.Values(ci).ToString()
                        End If
                        ' Escape quotes
                        val = val.Replace("""", "\""")
                        rowsArray.Append("""" & val & """")
                        If ci < agg.Schema.Count - 1 Then rowsArray.Append(",")
                    Next
                    rowsArray.Append("]")
                    If ri < groupRows.Count - 1 Then rowsArray.Append(",")
                Next
                rowsArray.Append("]}")

                Dim systemTemplate = context.SP_MergeDateRows
                Dim systemPrompt = interpolateSystemPromptFunc(systemTemplate)

                Dim userText =
                "MERGE_INSTRUCTION: " & If(String.IsNullOrWhiteSpace(mergeInstruction), "(none)", mergeInstruction) & vbCrLf &
                "INPUT_GROUP_JSON:" & vbCrLf & rowsArray.ToString()

                Dim mergedRow As ExtractionRow = Nothing
                Try
                    Dim resp = Await llmFunc(systemPrompt, userText, "", "", 0, useSecondApi, False)
                    resp = WebAgentInterpreter.SanitizeLlmResult(resp)
                    If Not String.IsNullOrWhiteSpace(resp) Then
                        Dim jt = Newtonsoft.Json.Linq.JToken.Parse(resp)
                        Dim valsTok = jt("values")
                        If valsTok IsNot Nothing AndAlso valsTok.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                            Dim er As New ExtractionRow With {.Values = New System.Collections.Generic.List(Of Object)()}
                            For ci = 0 To agg.Schema.Count - 1
                                Dim vTok = valsTok(ci)
                                er.Values.Add(If(vTok Is Nothing, "", vTok.ToString()))
                            Next
                            mergedRow = er
                        End If
                    End If
                Catch ex As Exception
                    agg.Errors.Add("Merge LLM failed for date " & g.Key & ": " & ex.Message)
                End Try

                If mergedRow Is Nothing Then
                    mergedRow = FallbackMergeRows(groupRows, agg.Schema, dateIdx)
                End If
                newRows.Add(mergedRow)
            Next

            ' Replace rows: keep any rows that had empty/invalid date (not grouped) + merged grouped
            Dim ungroupped = agg.Rows.Where(Function(r) String.IsNullOrWhiteSpace(GetNormalizedDateValue(r, dateIdx))).ToList()
            agg.Rows = ungroupped.Concat(newRows).ToList()
        End Function

        Private Function GetNormalizedDateValue(row As ExtractionRow, dateColIdxZeroBased As Integer) As String
            If row Is Nothing OrElse dateColIdxZeroBased < 0 OrElse dateColIdxZeroBased >= row.Values.Count Then Return ""
            Dim raw = CStr(row.Values(dateColIdxZeroBased))
            If String.IsNullOrWhiteSpace(raw) Then Return ""
            Return NormalizeDate(raw)
        End Function

        Private Function FallbackMergeRows(rows As System.Collections.Generic.List(Of ExtractionRow),
                                           schema As System.Collections.Generic.List(Of ExtractionSchemaColumn),
                                           dateColIdxZeroBased As Integer) As ExtractionRow
            ' Simple heuristic: keep date, first non-empty for others; concatenate text fields.
            Dim merged As New ExtractionRow With {.Values = New System.Collections.Generic.List(Of Object)()}
            For i = 0 To schema.Count - 1
                If i = dateColIdxZeroBased Then
                    merged.Values.Add(GetNormalizedDateValue(rows(0), dateColIdxZeroBased))
                    Continue For
                End If
                Dim colType = schema(i).Type
                Dim collectedText As New System.Text.StringBuilder()
                Dim chosen As Object = ""
                For Each r In rows
                    If i < r.Values.Count Then
                        Dim v = r.Values(i)
                        If v IsNot Nothing Then
                            Dim s = v.ToString().Trim()
                            If s.Length > 0 Then
                                If colType = "number" OrElse colType = "date" OrElse colType = "datetime" Then
                                    chosen = s
                                    Exit For
                                Else
                                    If collectedText.Length > 0 Then collectedText.Append(" | ")
                                    collectedText.Append(s)
                                    If chosen Is Nothing OrElse CStr(chosen).Length = 0 Then chosen = s
                                End If
                            End If
                        End If
                    End If
                Next
                If colType = "text" OrElse colType = "other" Then
                    merged.Values.Add(If(collectedText.Length > 0, collectedText.ToString(), CStr(chosen)))
                Else
                    merged.Values.Add(chosen)
                End If
            Next
            Return merged
        End Function


    End Module
End Namespace