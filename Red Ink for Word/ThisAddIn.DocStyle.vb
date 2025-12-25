' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.DocStyle.vb
' Purpose: LLM-assisted document styling system with two-step workflow.
'
' STEP 1 - CREATE STYLE TEMPLATE (ExtractParagraphStylesToJson):
'   Creates a reusable style template from a Word document where each paragraph
'   serves as a style definition. The source document should contain paragraphs
'   formatted as "STYLE NAME: description of when to apply this style..."
'   
'   The template captures:
'   - User-defined style names and application rules ("whenToApply")
'   - Underlying Word style definitions (wdStyles) with full formatting
'   - Paragraph formatting overrides (indentation, spacing, alignment)
'   - Font formatting, list/numbering settings, tab stops, borders, shading
'   
'   Output: JSON file stored in central (shared) or local (personal) path.
'
' STEP 2 - APPLY STYLE TEMPLATE (ApplyStyleTemplate):
'   Applies a style template to a target document using LLM analysis.
'   
'   The LLM analyzes each paragraph's content and structure to determine
'   the most appropriate user style from the template based on:
'   - Semantic meaning and purpose of the paragraph
'   - Structural clues (lists, indentation, outline level)
'   - Document context provided by the user
'   
'   Two modes available:
'   - Fast Mode: Single LLM call creates mapping plan for all paragraphs
'   - Detailed Mode: Per-paragraph LLM calls with full formatting control
'   
'   Features: Track changes, preview mode, confidence thresholds,
'   list preservation, table cell processing, Word style creation/update.
'
' Architecture:
'   - DocStyleSettings: Persistent user preferences via Registry
'   - DocStyleTemplate: Template file metadata and path management
'   - Central/Local paths: Supports shared team templates and personal ones
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Dim INI_DocStylePath As String = "C:\Users\david\Desktop\"
    Dim INI_DocStylePathLocal As String = "%Appdata%\Microsoft\Word\"

    ''' <summary>
    ''' Cache for shared list templates during style application.
    ''' Key is the template fingerprint, value is the created ListTemplate.
    ''' </summary>
    Private _sharedListTemplates As New Dictionary(Of String, Word.ListTemplate)()

    ' Treat bullets as break? (set True if you want bullet->numbered to restart)
    Private Const RuleBreakOnBullets As Boolean = False
    Private Const IndentStepPoints As Single = 10.0F

    Private Const ReportParaPreviewLen As Integer = 50

    ' Default if setting is missing/uninitialized.
    Private Const DefaultRuleHeadingOutlineLevelMax As Integer = 4

#Region "DocStyle Constants and Prompts"

    ' System prompt for applying styles (detailed mode - paragraph by paragraph)
    Private Const SP_ApplyDocStyle As String = "You are an expert document formatting assistant. Your task is to analyze a paragraph and determine the appropriate user style to apply based on a style template." & vbCrLf &
        "You will receive:" & vbCrLf &
        "1. A <STYLETEMPLATE> containing user style definitions with 'whenToApply' descriptions explaining when each style should be used" & vbCrLf &
        "2. A <CURRENTPARA> containing the paragraph text to format" & vbCrLf &
        "3. A <CURRENTPROPS> containing the paragraph's current formatting - USE THIS to understand the paragraph's role:" & vbCrLf &
        "   - If 'listFormatting.hasList' is true, the paragraph has bullets/numbering that should typically be preserved" & vbCrLf &
        "   - Check 'listFormatting.listType' and 'listFormatting.listString' to understand the list structure" & vbCrLf &
        "   - Consider 'paragraphFormatting.outlineLevel' to identify heading levels" & vbCrLf &
        "   - Look at indentation patterns to understand document hierarchy" & vbCrLf &
        "4. A <DOCUMENT> containing the full document context" & vbCrLf &
        "5. Optionally, <CONTEXT> with additional hints about the document type" & vbCrLf & vbCrLf &
        "IMPORTANT: Match the user style based on BOTH:" & vbCrLf &
        "- The semantic meaning and purpose of <CURRENTPARA> within the document" & vbCrLf &
        "- The structural clues in <CURRENTPROPS> (lists, indentation, outline level)" & vbCrLf & vbCrLf &
        "For list items: If the source paragraph has bullets/numbering (listFormatting.hasList=true), set 'preserveListFormatting' to true to keep the original list format as an override after applying the style." & vbCrLf & vbCrLf &
        "RESPOND ONLY with a valid JSON object in this exact format:" & vbCrLf &
        "{" & vbCrLf &
        "  ""userStyleName"": ""<name of the user style to apply>"", " & vbCrLf &
        "  ""confidence"": <0-100 confidence score>," & vbCrLf &
        "  ""reasoning"": ""<brief explanation>"", " & vbCrLf &
        "  ""preserveListFormatting"": <true if source has list formatting that should be kept, false otherwise>," & vbCrLf &
        "  ""formatting"": {" & vbCrLf &
        "    ""alignment"": ""<wdAlignParagraphLeft|wdAlignParagraphCenter|wdAlignParagraphRight|wdAlignParagraphJustify or null>"", " & vbCrLf &
        "    ""leftIndent"": <points or null>," & vbCrLf &
        "    ""rightIndent"": <points or null>," & vbCrLf &
        "    ""firstLineIndent"": <points or null>," & vbCrLf &
        "    ""spaceBefore"": <points or null>," & vbCrLf &
        "    ""spaceAfter"": <points or null>," & vbCrLf &
        "    ""lineSpacing"": <points or null>," & vbCrLf &
        "    ""lineSpacingRule"": ""<wdLineSpaceSingle|wdLineSpace1pt5|wdLineSpaceDouble|etc or null>"", " & vbCrLf &
        "    ""fontName"": ""<font name or null>"", " & vbCrLf &
        "    ""fontSize"": <size or null>," & vbCrLf &
        "    ""bold"": <true|false|null>," & vbCrLf &
        "    ""italic"": <true|false|null>," & vbCrLf &
        "    ""underline"": ""<wdUnderlineNone|wdUnderlineSingle|etc or null>""" & vbCrLf &
        "  }" & vbCrLf &
        "}" & vbCrLf & vbCrLf &
        "Use null for properties that should remain unchanged. Only include formatting properties that should be explicitly set."

    ' System prompt for fast mode (styles only, one call for mapping plan)
    Private Const SP_ApplyDocStyleFast As String = "You are an expert document formatting assistant. Your task is to create a mapping plan that associates each paragraph in a document with the most appropriate user style from a template." & vbCrLf &
    "You will receive:" & vbCrLf &
    "1. A <STYLETEMPLATE> containing user style definitions with 'whenToApply' descriptions" & vbCrLf &
    "2. A <DOCUMENT> containing the full document with numbered paragraphs and formatting hints" & vbCrLf &
    "3. Optionally, <CONTEXT> with additional hints about the document type" & vbCrLf & vbCrLf &
    "Each paragraph in <DOCUMENT> is formatted as:" & vbCrLf &
    "[index] [LIST:type/level] [INDENT:left/first] text..." & vbCrLf & vbCrLf &
    "CRITICAL MATCHING RULES:" & vbCrLf &
    "- [LIST:Bullet/...] paragraphs MUST be matched to user styles with 'Bullet' in their whenToApply description" & vbCrLf &
    "- [LIST:SimpleNumbering/...] or [LIST:OutlineNumbering/...] paragraphs MUST be matched to user styles with 'Numbering' or 'Numbered' in their whenToApply description" & vbCrLf &
    "- Do NOT assign a numbered list style to a bullet paragraph or vice versa" & vbCrLf &
    "- INDENT values indicate hierarchy - larger indents often mean sub-items" & vbCrLf & vbCrLf &
    "RESPOND ONLY with a valid JSON array where each element represents a paragraph:" & vbCrLf &
    "[" & vbCrLf &
    "  {""paragraphIndex"": 1, ""userStyleName"": ""<user style name>"", ""confidence"": <0-100>, ""preserveList"": <true|false>, ""restartNumbering"": <true|false>}," & vbCrLf &
    "  {""paragraphIndex"": 2, ""userStyleName"": ""<user style name>"", ""confidence"": <0-100>, ""preserveList"": <true|false>, ""restartNumbering"": <true|false>}," & vbCrLf &
    "  ..." & vbCrLf &
    "]" & vbCrLf & vbCrLf &
    "Use the exact user style names from the template. Set preserveList to true if the paragraph has list formatting that should be kept. Set restartNumbering to true if this numbered list item should restart numbering at 1 (e.g., first item of a new list section after non-list paragraphs or a heading). If no style is appropriate, use ""Normal"" or the closest match."

    ' Extended prompt suffix for LLM-assisted numbering reset
    Private Const SP_ApplyDocStyleFast_NumberingHint As String = vbCrLf & vbCrLf &
    "NUMBERING RESTART RULES:" & vbCrLf &
    "- Set restartNumbering=true for the FIRST numbered list item after:" & vbCrLf &
    "  * A heading or title paragraph" & vbCrLf &
    "  * One or more non-list body paragraphs" & vbCrLf &
    "  * A different type of list (e.g., bullet list followed by numbered list)" & vbCrLf &
    "  * A clear semantic section break (new topic, new article, new clause)" & vbCrLf &
    "- Set restartNumbering=false for numbered items that continue an existing sequence" & vbCrLf &
    "- Bullet lists do not need restartNumbering"

    ' Persistent settings key prefix
    Private Const DocStyleSettingsPrefix As String = "DocStyle_"

#End Region

#Region "DocStyle Settings Class"

    ''' <summary>
    ''' Holds persistent settings for DocStyle operations.
    ''' Persisted via My.Settings only (no registry usage).
    ''' </summary>
    Private Class DocStyleSettings
        Public Property TrackChanges As Boolean = False
        Public Property ApplyStyleDefinitions As Boolean = True
        Public Property PreviewMode As Boolean = False
        Public Property ConfidenceThreshold As Integer = 70
        Public Property UseConfidenceThreshold As Boolean = False
        Public Property ProcessTables As Boolean = True
        Public Property ParagraphsPerCall As Integer = 1
        Public Property FastModeStylesOnly As Boolean = True
        Public Property ShowReport As Boolean = True
        Public Property DocumentContext As String = ""
        Public Property UseSecondaryModel As Boolean = False
        Public Property ListNumberingReset As Integer = 0 ' 0=Off, 1=Rule-based, 2=LLM-assisted

        ' Configurable "heading breaks numbering run" threshold for rule-based restart.
        Public Property RuleHeadingOutlineLevelMax As Integer = DefaultRuleHeadingOutlineLevelMax

        Private Shared Function TryReadSetting(Of T)(settingName As String, ByRef value As T) As Boolean
            Try
                If My.Settings Is Nothing Then Return False

                Dim o As Object = My.Settings(settingName)
                If o Is Nothing Then Return False

                ' Handle common loose typing cases from Settings designer (String/Object).
                If GetType(T) Is GetType(Boolean) Then
                    value = CType(CObj(System.Convert.ToBoolean(o)), T)
                    Return True
                End If

                If GetType(T) Is GetType(Integer) Then
                    value = CType(CObj(System.Convert.ToInt32(o)), T)
                    Return True
                End If

                If GetType(T) Is GetType(String) Then
                    value = CType(CObj(System.Convert.ToString(o)), T)
                    Return True
                End If

                If TypeOf o Is T Then
                    value = CType(o, T)
                    Return True
                End If

                ' Last resort
                value = CType(System.Convert.ChangeType(o, GetType(T)), T)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Sub TryWriteSetting(settingName As String, value As Object)
            Try
                If My.Settings Is Nothing Then Exit Sub
                My.Settings(settingName) = value
            Catch
                ' Ignore if setting doesn't exist yet or is readonly/misconfigured.
            End Try
        End Sub

        Public Sub Load()

            Dim b As Boolean
            Dim i As Integer
            Dim s As String

            If TryReadSetting("DocStyle_TrackChanges", b) Then TrackChanges = b
            If TryReadSetting("DocStyle_ApplyStyleDefinitions", b) Then ApplyStyleDefinitions = b
            If TryReadSetting("DocStyle_PreviewMode", b) Then PreviewMode = b
            If TryReadSetting("DocStyle_UseConfidenceThreshold", b) Then UseConfidenceThreshold = b
            If TryReadSetting("DocStyle_ProcessTables", b) Then ProcessTables = b
            If TryReadSetting("DocStyle_FastModeStylesOnly", b) Then FastModeStylesOnly = b
            If TryReadSetting("DocStyle_ShowReport", b) Then ShowReport = b
            If TryReadSetting("DocStyle_UseSecondaryModel", b) Then UseSecondaryModel = b

            If TryReadSetting("DocStyle_ConfidenceThreshold", i) Then ConfidenceThreshold = i
            If TryReadSetting("DocStyle_ParagraphsPerCall", i) Then ParagraphsPerCall = i
            If TryReadSetting("DocStyle_ListNumberingReset", i) Then ListNumberingReset = i

            If TryReadSetting("DocStyle_DocumentContext", s) Then DocumentContext = If(s, "")

            If TryReadSetting("DocStyle_RuleHeadingOutlineLevelMax", i) Then
                If i >= 0 AndAlso i <= 9 Then
                    RuleHeadingOutlineLevelMax = i
                End If
            End If

            ' Clamp defense (also protects defaults if settings are corrupted).
            If ConfidenceThreshold < 0 Then ConfidenceThreshold = 0
            If ConfidenceThreshold > 100 Then ConfidenceThreshold = 100

            If ParagraphsPerCall < 1 Then ParagraphsPerCall = 1
            If ListNumberingReset < 0 Then ListNumberingReset = 0
            If ListNumberingReset > 2 Then ListNumberingReset = 2

            If RuleHeadingOutlineLevelMax < 0 Then RuleHeadingOutlineLevelMax = 0
            If RuleHeadingOutlineLevelMax > 9 Then RuleHeadingOutlineLevelMax = 9
        End Sub

        Public Sub Save()
            TryWriteSetting("DocStyle_TrackChanges", TrackChanges)
            TryWriteSetting("DocStyle_ApplyStyleDefinitions", ApplyStyleDefinitions)
            TryWriteSetting("DocStyle_PreviewMode", PreviewMode)
            TryWriteSetting("DocStyle_ConfidenceThreshold", ConfidenceThreshold)
            TryWriteSetting("DocStyle_UseConfidenceThreshold", UseConfidenceThreshold)
            TryWriteSetting("DocStyle_ProcessTables", ProcessTables)
            TryWriteSetting("DocStyle_ParagraphsPerCall", ParagraphsPerCall)
            TryWriteSetting("DocStyle_FastModeStylesOnly", FastModeStylesOnly)
            TryWriteSetting("DocStyle_ShowReport", ShowReport)
            TryWriteSetting("DocStyle_DocumentContext", If(DocumentContext, ""))
            TryWriteSetting("DocStyle_UseSecondaryModel", UseSecondaryModel)
            TryWriteSetting("DocStyle_ListNumberingReset", ListNumberingReset)
            TryWriteSetting("DocStyle_RuleHeadingOutlineLevelMax", RuleHeadingOutlineLevelMax)

            Try
                If My.Settings IsNot Nothing Then
                    My.Settings.Save()
                End If
            Catch
            End Try
        End Sub
    End Class

#End Region

#Region "Step 1: Extract Paragraph Styles to JSON (Style Template Creation)"

    ''' <summary>
    ''' Extracts paragraph formatting from the selection (or entire document) and generates
    ''' a comprehensive JSON structure describing each paragraph's text, style, and formatting.
    ''' The JSON is designed to be consumed by an LLM to apply formatting to other documents.
    ''' </summary>
    Public Sub ExtractParagraphStylesToJson()
        Try
            Dim app As Word.Application = Globals.ThisAddIn.Application
            Dim doc As Word.Document = app.ActiveDocument

            If doc Is Nothing Then
                ShowCustomMessageBox("No active document found.")
                Return
            End If

            ' Expand paths
            Dim docStylePath As String = ExpandEnvironmentVariables(INI_DocStylePath)
            If Not String.IsNullOrEmpty(docStylePath) AndAlso Not docStylePath.EndsWith("\") Then
                docStylePath &= "\"
            End If

            Dim docStylePathLocal As String = ExpandEnvironmentVariables(INI_DocStylePathLocal)
            If Not String.IsNullOrEmpty(docStylePathLocal) AndAlso Not docStylePathLocal.EndsWith("\") Then
                docStylePathLocal &= "\"
            End If

            Dim hasGlobal As Boolean = Not String.IsNullOrWhiteSpace(docStylePath) AndAlso Directory.Exists(docStylePath)
            Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(docStylePathLocal) AndAlso Directory.Exists(docStylePathLocal)

            ' Determine save path
            Dim savePath As String
            Dim isLocal As Boolean = False

            If Not hasGlobal AndAlso Not hasLocal Then
                ' Use desktop and warn
                savePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                ShowCustomMessageBox("Warning: Neither 'DocStylePath' nor 'DocStylePathLocal' is configured. The file will be saved to your Desktop.")
            ElseIf hasGlobal AndAlso hasLocal Then
                ' Ask user
                Dim choice As Integer = ShowCustomYesNoBox("Where do you want to save the style template?",
                    "Central (shared)", "Local (personal)", $"{AN} - Save Location")
                If choice = 0 Then Return
                If choice = 1 Then
                    savePath = docStylePath
                    isLocal = False
                Else
                    savePath = docStylePathLocal
                    isLocal = True
                End If
            ElseIf hasGlobal Then
                savePath = docStylePath
            Else
                savePath = docStylePathLocal
                isLocal = True
            End If

            ' Ask for template display name using ShowCustomInputBox
            Dim safeDocName As String = Regex.Replace(doc.Name.Replace(".docx", "").Replace(".doc", ""), "[^a-zA-Z0-9_-]", "_")
            Dim defaultDisplayName As String = $"{safeDocName}_{DateTime.Now:yyyyMMdd_HHmm}"

            Dim templateDisplayName As String = ShowCustomInputBox("Enter a name for this style template:", $"{AN} - Style Template Name", True, defaultDisplayName)
            If String.IsNullOrWhiteSpace(templateDisplayName) OrElse templateDisplayName.Equals("ESC", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            Dim targetRange As Word.Range

            ' Use selection if available, otherwise entire document
            If app.Selection.Type = WdSelectionType.wdSelectionIP Then
                targetRange = doc.Content
            Else
                targetRange = app.Selection.Range.Duplicate
            End If

            Dim userStyles As New JArray()
            Dim wdStyleDefinitions As New JObject()
            Dim collectedStyles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim processedCount As Integer = 0
            Dim errorCount As Integer = 0
            Dim unparsedStyleNameCount As Integer = 0

            ' Process all paragraphs including those in tables
            For Each para As Word.Paragraph In targetRange.Paragraphs
                Try
                    Dim parseResult As (userStyleJson As JObject, parsedStyleName As Boolean) = ExtractUserStyleFromParagraphEx(para, processedCount + 1)
                    If parseResult.userStyleJson IsNot Nothing Then
                        userStyles.Add(parseResult.userStyleJson)
                        processedCount += 1

                        If Not parseResult.parsedStyleName Then
                            unparsedStyleNameCount += 1
                        End If

                        ' Collect wdStyle name for style definitions
                        Dim wdStyleName As String = If(parseResult.userStyleJson("wdStyleName") IsNot Nothing, parseResult.userStyleJson("wdStyleName").ToString(), "")
                        If Not String.IsNullOrWhiteSpace(wdStyleName) AndAlso Not collectedStyles.Contains(wdStyleName) Then
                            collectedStyles.Add(wdStyleName)
                        End If
                    End If
                Catch ex As Exception
                    errorCount += 1
                    Debug.WriteLine($"Error processing paragraph {processedCount + 1}: {ex.Message}")
                End Try
            Next

            ' Extract full wdStyle definitions for collected styles
            For Each styleName In collectedStyles
                Try
                    Dim styleObj As JObject = ExtractFullStyleDefinition(doc, styleName)
                    If styleObj IsNot Nothing Then
                        wdStyleDefinitions(styleName) = styleObj
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error extracting wdStyle definition for '{styleName}': {ex.Message}")
                End Try
            Next

            ' Build the complete JSON structure
            Dim result As New JObject()
            result("templateName") = templateDisplayName
            result("description") = "Style template for intelligent document formatting. Each user style in 'userStyles' includes a 'whenToApply' field describing the situations where that style should be applied. The 'wdStyleDefinitions' section contains Word style definitions that can optionally be created/updated in target documents."
            result("documentInfo") = New JObject From {
                {"extractedFrom", doc.Name},
                {"extractionDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")},
                {"totalUserStyles", processedCount},
                {"totalWdStyles", collectedStyles.Count}
            }
            result("userStyles") = userStyles
            result("wdStyleDefinitions") = wdStyleDefinitions

            ' Serialize to formatted JSON
            Dim jsonString As String = JsonConvert.SerializeObject(result, Formatting.Indented)

            ' Generate filename using safe template name
            Dim safeTemplateName As String = Regex.Replace(templateDisplayName, "[^a-zA-Z0-9_-]", "_")
            Dim fileName As String = $"{AN2}-ds-{safeTemplateName}.json"
            Dim filePath As String = Path.Combine(savePath, fileName)

            ' Check if file exists and handle overwrite
            If File.Exists(filePath) Then
                Dim overwrite As Integer = ShowCustomYesNoBox($"A style template with this name already exists.{vbCrLf}{vbCrLf}Do you want to overwrite it?",
                    "Overwrite", "Cancel", $"{AN} - File Exists")
                If overwrite <> 1 Then Return
            End If

            Try
                File.WriteAllText(filePath, jsonString, Encoding.UTF8)
            Catch ioEx As Exception
                ShowCustomMessageBox($"Could not save file: {ioEx.Message}")
                Return
            End Try

            ' Build success message - only show tip if some styles were not parsed
            Dim successMessage As String = $"Style template '{templateDisplayName}' has been created successfully.{vbCrLf}{vbCrLf}Location: {filePath}{vbCrLf}{vbCrLf}Would you like to edit the template now?"

            If unparsedStyleNameCount > 0 Then
                successMessage &= $"{vbCrLf}{vbCrLf}Tip: {unparsedStyleNameCount} style(s) could not be auto-named. Edit the 'userStyleName' and 'whenToApply' fields for these styles. Use format 'STYLE NAME: description...' in your source paragraphs for automatic parsing."
            End If

            Dim editChoice As Integer = ShowCustomYesNoBox(successMessage,
                "Edit Template", "Close", $"{AN} - Style Template Created")

            If editChoice = 1 Then
                SLib.ShowTextFileEditor(filePath, $"{AN} - Style Template '{templateDisplayName}'", True, _context)
            End If

        Catch ex As Exception
            ShowCustomMessageBox($"Error extracting paragraph styles: {ex.Message}")
        End Try
    End Sub


    Private Shared Function NormalizeStyleKey(s As String) As String
        If s Is Nothing Then Return ""

        ' Normalize common invisible differences from LLM / copy-paste.
        s = s.Replace(ChrW(&HA0), " ") ' NBSP -> space
        s = Regex.Replace(s, "\s+", " ").Trim()

        Return s
    End Function

    ''' <summary>
    ''' Extracts user style information from a single paragraph for template creation.
    ''' Parses the paragraph text to extract userStyleName and whenToApply.
    ''' Returns a tuple with the JObject and a boolean indicating if the style name was successfully parsed.
    ''' </summary>
    Private Function ExtractUserStyleFromParagraphEx(para As Word.Paragraph, index As Integer) As (userStyleJson As JObject, parsedStyleName As Boolean)
        Dim paraRange As Word.Range = para.Range.Duplicate

        ' Skip empty paragraphs (only paragraph mark)
        Dim text As String = paraRange.Text
        If String.IsNullOrWhiteSpace(text) OrElse text = vbCr OrElse text = vbCrLf Then
            Return (Nothing, False)
        End If

        ' Remove trailing paragraph mark for cleaner text
        text = text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10)).Trim()

        Dim result As New JObject()
        result("userStyleIndex") = index

        ' Try to parse "STYLE NAME: description..." format
        Dim parsedStyleName As Boolean = False
        Dim userStyleName As String = $"UserStyle_{index}"
        Dim whenToApply As String = text

        ' Look for colon separator - style name should be short (max ~50 chars) and before colon
        Dim colonIndex As Integer = text.IndexOf(":"c)
        If colonIndex > 0 AndAlso colonIndex <= 50 Then
            Dim potentialName As String = text.Substring(0, colonIndex).Trim()
            Dim potentialDescription As String = text.Substring(colonIndex + 1).Trim()

            ' Validate: name should not contain line breaks and should have some description after
            If Not String.IsNullOrWhiteSpace(potentialName) AndAlso
               Not potentialName.Contains(vbCr) AndAlso
               Not potentialName.Contains(vbLf) AndAlso
               Not String.IsNullOrWhiteSpace(potentialDescription) Then
                userStyleName = potentialName
                whenToApply = potentialDescription
                parsedStyleName = True
            End If
        End If

        result("userStyleName") = userStyleName
        result("whenToApply") = whenToApply

        ' Get the Word style name
        Try
            Dim style As Word.Style = para.Style
            result("wdStyleName") = style.NameLocal
            result("wdStyleBuiltIn") = style.BuiltIn
        Catch
            result("wdStyleName") = "Normal"
            result("wdStyleBuiltIn") = True
        End Try

        ' Check if in table
        Dim isInTable As Boolean = False
        Try
            If paraRange.Tables.Count > 0 OrElse paraRange.Cells.Count > 0 Then
                isInTable = True
            End If
        Catch
        End Try
        result("isInTableCell") = isInTable

        ' Paragraph formatting
        result("paragraphFormatting") = ExtractParagraphFormat(para, paraRange)

        ' Font/Character formatting
        result("fontFormatting") = ExtractFontFormat(paraRange)

        ' List formatting
        result("listFormatting") = ExtractListFormat(paraRange)

        ' Tab stops (compact format)
        result("tabStops") = ExtractTabStopsCompact(para)

        ' Borders (only if present)
        Dim borders As JObject = ExtractBorders(para)
        If borders.Properties().Any(Function(p) p.Name <> "distanceFromText" AndAlso p.Name <> "error") Then
            result("borders") = borders
        End If

        ' Shading (only if non-default)
        Dim shading As JObject = ExtractShading(para)
        If shading("backgroundColor") IsNot Nothing OrElse shading("foregroundColor") IsNot Nothing Then
            result("shading") = shading
        End If

        Return (result, parsedStyleName)
    End Function

    ''' <summary>
    ''' Extracts tab stops in a compact format for efficiency.
    ''' </summary>
    Private Function ExtractTabStopsCompact(para As Word.Paragraph) As JArray
        Dim tabs As New JArray()
        Try
            For Each tabStop As Word.TabStop In para.TabStops
                Dim tab As New JObject()
                tab("pos") = Math.Round(tabStop.Position, 2)
                tab("align") = tabStop.Alignment.ToString().Replace("wdAlignTab", "")
                If tabStop.Leader <> WdTabLeader.wdTabLeaderSpaces Then
                    tab("leader") = tabStop.Leader.ToString().Replace("wdTabLeader", "")
                End If
                tabs.Add(tab)
            Next
        Catch
        End Try
        Return tabs
    End Function

#End Region

    ''' <summary>
    ''' Extracts complete wdStyle definition including base style hierarchy, tab stops,
    ''' list template with all levels, bullet characters, and linked level information.
    ''' </summary>
    Private Function ExtractFullStyleDefinition(doc As Word.Document, styleName As String) As JObject
        Try
            Dim style As Word.Style = doc.Styles(styleName)
            If style Is Nothing Then Return Nothing

            Dim styleDef As New JObject()
            styleDef("wdStyleName") = style.NameLocal
            styleDef("styleType") = style.Type.ToString()
            styleDef("builtIn") = style.BuiltIn

            ' Base style hierarchy
            Dim baseStyles As New JArray()
            Dim currentStyle As Word.Style = style
            Try
                While currentStyle.BaseStyle IsNot Nothing
                    Dim baseStyle As Word.Style = CType(currentStyle.BaseStyle, Word.Style)
                    baseStyles.Add(baseStyle.NameLocal)
                    currentStyle = baseStyle
                End While
            Catch
            End Try
            If baseStyles.Count > 0 Then
                styleDef("baseStyleHierarchy") = baseStyles
            End If

            ' Next paragraph style
            Try
                If style.NextParagraphStyle IsNot Nothing Then
                    styleDef("nextParagraphStyle") = CType(style.NextParagraphStyle, Word.Style).NameLocal
                End If
            Catch
            End Try

            ' Paragraph formatting from style
            Try
                Dim pf As Word.ParagraphFormat = style.ParagraphFormat
                Dim paraFormat As New JObject()
                paraFormat("alignment") = pf.Alignment.ToString()
                paraFormat("leftIndent") = pf.LeftIndent
                paraFormat("rightIndent") = pf.RightIndent
                paraFormat("firstLineIndent") = pf.FirstLineIndent
                paraFormat("spaceBefore") = pf.SpaceBefore
                paraFormat("spaceAfter") = pf.SpaceAfter
                paraFormat("lineSpacing") = pf.LineSpacing
                paraFormat("lineSpacingRule") = pf.LineSpacingRule.ToString()
                paraFormat("keepTogether") = pf.KeepTogether
                paraFormat("keepWithNext") = pf.KeepWithNext
                paraFormat("pageBreakBefore") = pf.PageBreakBefore
                paraFormat("widowControl") = pf.WidowControl
                paraFormat("outlineLevel") = pf.OutlineLevel.ToString()
                styleDef("paragraphFormat") = paraFormat
            Catch
            End Try

            ' Tab stops from style
            Try
                Dim tabs As New JArray()
                For Each tabStop As Word.TabStop In style.ParagraphFormat.TabStops
                    Dim tab As New JObject()
                    tab("position") = Math.Round(tabStop.Position, 2)
                    tab("alignment") = tabStop.Alignment.ToString()
                    tab("leader") = tabStop.Leader.ToString()
                    tabs.Add(tab)
                Next
                If tabs.Count > 0 Then
                    styleDef("tabStops") = tabs
                End If
            Catch
            End Try

            ' Font formatting from style - COMPLETE extraction
            Try
                Dim font As Word.Font = style.Font
                Dim fontFormat As New JObject()
                fontFormat("name") = font.Name
                fontFormat("size") = font.Size
                fontFormat("bold") = (font.Bold = -1)
                fontFormat("italic") = (font.Italic = -1)
                fontFormat("underline") = font.Underline.ToString()
                fontFormat("allCaps") = (font.AllCaps = -1)
                fontFormat("smallCaps") = (font.SmallCaps = -1)
                fontFormat("strikeThrough") = (font.StrikeThrough = -1)
                fontFormat("doubleStrikeThrough") = (font.DoubleStrikeThrough = -1)
                fontFormat("subscript") = (font.Subscript = -1)
                fontFormat("superscript") = (font.Superscript = -1)
                fontFormat("color") = font.Color.ToString()
                fontFormat("colorRGB") = ColorToRGB(font.Color)
                Try
                    fontFormat("scaling") = font.Scaling
                    fontFormat("spacing") = font.Spacing
                    fontFormat("position") = font.Position
                    fontFormat("kerning") = font.Kerning
                Catch
                End Try
                styleDef("fontFormat") = fontFormat
            Catch
            End Try

            ' List formatting from style (linked list template) - COMPLETE extraction
            Try
                Dim listTemplate As Word.ListTemplate = style.ListTemplate
                If listTemplate IsNot Nothing Then
                    Dim listFormat As New JObject()
                    listFormat("hasListTemplate") = True
                    listFormat("outlineNumbered") = listTemplate.OutlineNumbered

                    ' CRITICAL: Store which level this style is linked to
                    Try
                        listFormat("linkedLevel") = style.ListLevelNumber
                    Catch
                        listFormat("linkedLevel") = 1
                    End Try

                    ' Generate a fingerprint to identify shared templates
                    ' Styles sharing the same outline list template will have identical fingerprints
                    Try
                        Dim fingerprint As New StringBuilder()
                        For lvl As Integer = 1 To Math.Min(9, listTemplate.ListLevels.Count)
                            Try
                                Dim level As Word.ListLevel = listTemplate.ListLevels(lvl)
                                fingerprint.Append($"{lvl}:{level.NumberStyle}:{level.NumberFormat}|")
                            Catch
                            End Try
                        Next
                        listFormat("templateFingerprint") = fingerprint.ToString()
                    Catch
                    End Try

                    ' Extract ALL levels with complete information
                    Dim levels As New JArray()
                    For levelNum As Integer = 1 To listTemplate.ListLevels.Count
                        Try
                            Dim level As Word.ListLevel = listTemplate.ListLevels(levelNum)
                            Dim levelInfo As New JObject()
                            levelInfo("level") = levelNum
                            levelInfo("numberStyle") = level.NumberStyle.ToString()
                            levelInfo("textPosition") = level.TextPosition
                            levelInfo("tabPosition") = level.TabPosition
                            levelInfo("numberPosition") = level.NumberPosition
                            levelInfo("alignment") = level.Alignment.ToString()
                            levelInfo("startAt") = level.StartAt

                            ' Handle bullet styles - capture the actual character and font
                            If level.NumberStyle = WdListNumberStyle.wdListNumberStyleBullet Then
                                Try
                                    ' Get bullet font - check multiple ways
                                    Dim bulletFontName As String = ""
                                    Try
                                        If level.Font IsNot Nothing AndAlso Not String.IsNullOrEmpty(level.Font.Name) Then
                                            bulletFontName = level.Font.Name
                                        End If
                                    Catch
                                    End Try

                                    If Not String.IsNullOrEmpty(bulletFontName) Then
                                        levelInfo("bulletFont") = bulletFontName
                                    Else
                                        ' Common bullet fonts as fallback detection
                                        levelInfo("bulletFont") = "Symbol"
                                    End If

                                    ' Get bullet character as Unicode code point
                                    If Not String.IsNullOrEmpty(level.NumberFormat) AndAlso level.NumberFormat.Length > 0 Then
                                        Dim bulletChar As Char = level.NumberFormat.Chars(0)
                                        levelInfo("bulletCharCode") = AscW(bulletChar)
                                        ' Store original for debugging but don't rely on it
                                        levelInfo("numberFormat") = ""
                                    Else
                                        ' Default bullet character code (standard bullet •)
                                        levelInfo("bulletCharCode") = &H2022
                                        levelInfo("numberFormat") = ""
                                    End If
                                Catch ex As Exception
                                    Debug.WriteLine($"Error extracting bullet info for level {levelNum}: {ex.Message}")
                                    levelInfo("bulletCharCode") = &H2022
                                    levelInfo("bulletFont") = "Symbol"
                                    levelInfo("numberFormat") = ""
                                End Try
                            Else
                                ' Non-bullet: store numberFormat as-is
                                levelInfo("numberFormat") = level.NumberFormat
                            End If

                            Try
                                levelInfo("trailingCharacter") = level.TrailingCharacter.ToString()
                            Catch
                            End Try

                            levels.Add(levelInfo)
                        Catch
                        End Try
                    Next
                    listFormat("levels") = levels
                    styleDef("listFormat") = listFormat
                End If
            Catch
                ' Style has no list template
            End Try

            Return styleDef

        Catch ex As Exception
            Debug.WriteLine($"Error extracting style definition for '{styleName}': {ex.Message}")
            Return Nothing
        End Try
    End Function


#Region "Step 2: Apply Style Template"

    ''' <summary>
    ''' Main entry point for applying a style template to the document or selection.
    ''' </summary>
    Public Async Sub ApplyStyleTemplate()
        If INILoadFail() Then Return

        Dim do2ndModel As Boolean = False
        Dim settings As New DocStyleSettings()
        settings.Load()

        Try
            ' Expand paths
            Dim docStylePath As String = ExpandEnvironmentVariables(INI_DocStylePath)
            If Not String.IsNullOrEmpty(docStylePath) AndAlso Not docStylePath.EndsWith("\") Then
                docStylePath &= "\"
            End If

            Dim docStylePathLocal As String = ExpandEnvironmentVariables(INI_DocStylePathLocal)
            If Not String.IsNullOrEmpty(docStylePathLocal) AndAlso Not docStylePathLocal.EndsWith("\") Then
                docStylePathLocal &= "\"
            End If

            Dim hasGlobal As Boolean = Not String.IsNullOrWhiteSpace(docStylePath) AndAlso Directory.Exists(docStylePath)
            Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(docStylePathLocal) AndAlso Directory.Exists(docStylePathLocal)

            If Not hasGlobal AndAlso Not hasLocal Then
                ShowCustomMessageBox("No style template paths are configured. Please configure 'DocStylePath' or 'DocStylePathLocal' in your INI file.")
                Return
            End If

            ' Get Word application and document
            Dim app As Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
                ShowCustomMessageBox("No open document.")
                Return
            End If

            Dim doc As Word.Document = app.ActiveDocument
            If doc Is Nothing Then
                ShowCustomMessageBox("Active document was not found.")
                Return
            End If

            Dim currentSelection As Word.Selection = app.Selection
            Dim targetRange As Word.Range

            ' Check selection
            If currentSelection.Type = WdSelectionType.wdSelectionIP Then
                Dim answer As Integer = ShowCustomYesNoBox("You have not selected any text. Do you want to apply the style template to the entire document?",
                    "Yes, entire document", "No, cancel", $"{AN} - Apply Style Template")
                If answer <> 1 Then Return
                targetRange = doc.Content
            Else
                targetRange = currentSelection.Range.Duplicate
            End If

            ' Load available style templates
            Dim templates As List(Of DocStyleTemplate) = LoadStyleTemplates(docStylePath, docStylePathLocal)
            If templates Is Nothing OrElse templates.Count = 0 Then
                ShowCustomMessageBox($"No valid style templates found. Create templates using 'Extract Style Template' and save them as '{AN2}-ds-*.json' files.")
                Return
            End If

            ' Build display options using template display names
            Dim displayToTemplate As New Dictionary(Of String, DocStyleTemplate)(StringComparer.OrdinalIgnoreCase)
            Dim displayOptions As New List(Of String)()
            For Each t In templates
                Dim display As String = t.DisplayName
                If t.IsLocal Then display &= " (local)"
                ' Handle duplicate display names by appending index
                Dim originalDisplay As String = display
                Dim counter As Integer = 1
                While displayToTemplate.ContainsKey(display)
                    counter += 1
                    display = $"{originalDisplay} ({counter})"
                End While
                displayToTemplate(display) = t
                displayOptions.Add(display)
            Next

            ' Parameter form
            Dim p0 As New SLib.InputParameter("Style Template", If(displayOptions.Count > 0, displayOptions(0), ""))
            p0.Options = displayOptions

            Dim p1 As New SLib.InputParameter("Apply in Track Changes", settings.TrackChanges)
            Dim p2 As New SLib.InputParameter("Create/update Word styles from template", settings.ApplyStyleDefinitions)
            Dim p3 As New SLib.InputParameter("Preview each change", settings.PreviewMode)
            Dim p4 As New SLib.InputParameter("Use confidence threshold", settings.UseConfidenceThreshold)
            Dim p5 As New SLib.InputParameter("Confidence threshold (0-100%)", settings.ConfidenceThreshold)
            Dim p6 As New SLib.InputParameter("Process table cells", settings.ProcessTables)
            Dim p7 As New SLib.InputParameter("Apply only style, no text updating (faster, safer)", settings.FastModeStylesOnly)
            Dim p8 As New SLib.InputParameter("Paragraphs per LLM call (n/a in 'faster' mode)", settings.ParagraphsPerCall)

            ' List numbering reset option
            Dim listResetOptions As New List(Of String) From {"Off", "Rule-based (after non-list paragraphs)", "LLM-assisted (semantic analysis)"}
            Dim p8b As New SLib.InputParameter("List numbering reset", listResetOptions(settings.ListNumberingReset))
            p8b.Options = listResetOptions

            ' NEW: heading outline level max (only relevant for rule-based restart)
            Dim headingOutlineLevelOptions As New List(Of String) From {
                "0 (never treat outline levels as headings)",
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9"
            }
            Dim initialHeadingMax As Integer = settings.RuleHeadingOutlineLevelMax
            If initialHeadingMax < 0 Then initialHeadingMax = 0
            If initialHeadingMax > 9 Then initialHeadingMax = 9

            Dim p8c As New SLib.InputParameter("Maximum heading levels (rule-based reset)", headingOutlineLevelOptions(initialHeadingMax))
            p8c.Options = headingOutlineLevelOptions

            Dim p9 As New SLib.InputParameter("Show report at end", settings.ShowReport)
            Dim p10 As New SLib.InputParameter("Document context (optional)", settings.DocumentContext)

            Dim p11 As SLib.InputParameter
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                p11 = New SLib.InputParameter("Use a secondary model", settings.UseSecondaryModel)
            ElseIf INI_SecondAPI Then
                p11 = New SLib.InputParameter("Use the secondary model", settings.UseSecondaryModel)
            Else
                p11 = New SLib.InputParameter("Use the secondary model", CType(Nothing, Boolean?))
            End If

            ' CHANGED: inserted p8c after list reset parameter (p8b)
            Dim params() As SLib.InputParameter = {p0, p1, p2, p3, p4, p5, p6, p7, p8, p8b, p8c, p9, p10, p11}

            If Not ShowCustomVariableInputForm("Configure Style Template Application:", $"{AN} - Apply Style Template", params) Then
                Return
            End If

            ' Read back values
            Dim chosenDisplay As String = System.Convert.ToString(params(0).Value)
            settings.TrackChanges = System.Convert.ToBoolean(params(1).Value)
            settings.ApplyStyleDefinitions = System.Convert.ToBoolean(params(2).Value)
            settings.PreviewMode = System.Convert.ToBoolean(params(3).Value)
            settings.UseConfidenceThreshold = System.Convert.ToBoolean(params(4).Value)
            settings.ConfidenceThreshold = System.Convert.ToInt32(params(5).Value)
            settings.ProcessTables = System.Convert.ToBoolean(params(6).Value)
            settings.FastModeStylesOnly = System.Convert.ToBoolean(params(7).Value)
            settings.ParagraphsPerCall = Math.Max(1, System.Convert.ToInt32(params(8).Value))
            settings.ListNumberingReset = listResetOptions.IndexOf(System.Convert.ToString(params(9).Value))
            If settings.ListNumberingReset < 0 Then settings.ListNumberingReset = 0

            ' NEW: parse RuleHeadingOutlineLevelMax from p8c
            Dim headingMaxRaw As String = System.Convert.ToString(params(10).Value)
            Dim m As Match = Regex.Match(headingMaxRaw, "^\s*(\d+)")
            If m.Success Then
                settings.RuleHeadingOutlineLevelMax = Math.Max(0, Math.Min(9, Integer.Parse(m.Groups(1).Value)))
            Else
                settings.RuleHeadingOutlineLevelMax = DefaultRuleHeadingOutlineLevelMax
            End If

            settings.ShowReport = System.Convert.ToBoolean(params(11).Value)
            settings.DocumentContext = System.Convert.ToString(params(12).Value)

            Dim secondModel = params(13).Value
            If TypeOf secondModel Is Boolean Then
                do2ndModel = CBool(secondModel)
                settings.UseSecondaryModel = do2ndModel
            End If

            ' Save settings
            settings.Save()

            ' Resolve selected template
            Dim chosenTemplate As DocStyleTemplate = Nothing
            If Not displayToTemplate.TryGetValue(chosenDisplay, chosenTemplate) Then
                ShowCustomMessageBox("Selected style template could not be resolved.")
                Return
            End If

            ' Load template content
            Dim templateJson As String
            Try
                templateJson = File.ReadAllText(chosenTemplate.FilePath, Encoding.UTF8)
            Catch ex As Exception
                ShowCustomMessageBox($"Could not read template file: {ex.Message}")
                Return
            End Try

            Dim templateObj As JObject
            Try
                templateObj = JObject.Parse(templateJson)
            Catch ex As Exception
                ShowCustomMessageBox($"Invalid JSON in template file: {ex.Message}")
                Return
            End Try

            ' Handle secondary model
            If do2ndModel Then
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The secondary model could not be loaded - aborting.")
                        Return
                    End If
                End If
            End If

            ' Extract and apply wdStyle definitions if requested
            Dim wdStyleDefinitions As JObject = Nothing
            If templateObj("wdStyleDefinitions") IsNot Nothing Then
                wdStyleDefinitions = CType(templateObj("wdStyleDefinitions"), JObject)
                If settings.ApplyStyleDefinitions AndAlso wdStyleDefinitions IsNot Nothing Then
                    ApplyWdStyleDefinitionsToDocument(doc, wdStyleDefinitions)
                End If
                templateObj.Remove("wdStyleDefinitions")
            End If

            ' Create minimal template for LLM (only names and whenToApply)
            Dim templateForLLM As String = CreateMinimalTemplateForLLM(templateObj)

            ' Handle track changes
            Dim originalTrackChanges As Boolean = doc.TrackRevisions
            If settings.TrackChanges Then
                doc.TrackRevisions = True
            End If

            Try
                Using New WordUndoScope(app, $"{AN} - Apply Style Template")
                    ' Run the appropriate application mode
                    Dim report As String
                    If settings.FastModeStylesOnly Then
                        report = Await ApplyStylesFastMode(doc, targetRange, templateForLLM, templateObj, settings, do2ndModel)
                    Else
                        report = Await ApplyStylesDetailedMode(doc, targetRange, templateForLLM, templateObj, settings, do2ndModel)
                    End If

                    ' Show report if requested
                    If settings.ShowReport AndAlso Not String.IsNullOrWhiteSpace(report) Then
                        Dim reportResult As String = ShowCustomWindow(
                            "Style Template Application Complete",
                            report,
                            "The report has been copied to your clipboard.",
                            $"{AN} - Application Report",
                            NoRTF:=True,
                            Getfocus:=True)
                        SLib.PutInClipboard(report)
                    End If
                End Using
            Finally
                ' Restore track changes
                doc.TrackRevisions = originalTrackChanges
            End Try

        Catch ex As Exception
            ShowCustomMessageBox($"Error applying style template: {ex.Message}")
        Finally
            If do2ndModel AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub



    ' =========================
    ' DROP-IN REPLACEMENT: ApplyStylesFastMode
    '
    ' Fix:
    ' - Normalizes userStyleName keys (LLM + template) so invisible chars (NBSP/ZWSP/whitespace) can’t break lookup.
    ' - Uses dictionaries for O(1) lookups instead of scanning userStyles array per paragraph.
    ' - Logs when a userStyleName can’t be resolved to a template userStyleDef (so indent overrides don’t silently disappear).
    ' =========================
    Private Async Function ApplyStylesFastMode(doc As Word.Document, targetRange As Word.Range,
                                            templateJson As String, templateObj As JObject,
                                            settings As DocStyleSettings,
                                            useSecondAPI As Boolean) As Task(Of String)
        Dim report As New StringBuilder()
        report.AppendLine("=== Style Template Application Report (Fast Mode) ===")
        report.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        report.AppendLine()

        Try
            ' Build user style name -> wdStyle name mapping and user style name -> userStyleDef.
            Dim userStyleToWdStyle As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim userStyleNameToDef As New Dictionary(Of String, JObject)(StringComparer.OrdinalIgnoreCase)

            If templateObj("userStyles") IsNot Nothing Then
                For Each userStyle As JObject In CType(templateObj("userStyles"), JArray)
                    Dim rawUserStyleName As String = If(userStyle("userStyleName") IsNot Nothing, userStyle("userStyleName").ToString(), "")
                    Dim userStyleKey As String = NormalizeStyleKey(rawUserStyleName)

                    If String.IsNullOrWhiteSpace(userStyleKey) Then Continue For

                    Dim wdStyleName As String = If(userStyle("wdStyleName") IsNot Nothing, userStyle("wdStyleName").ToString(), "Normal")

                    userStyleToWdStyle(userStyleKey) = wdStyleName
                    userStyleNameToDef(userStyleKey) = userStyle
                Next
            End If

            ' Build numbered paragraph list with formatting hints
            Dim paragraphTexts As New StringBuilder()
            Dim paragraphList As New List(Of Word.Paragraph)()
            Dim paragraphHasList As New List(Of Boolean)()
            Dim idx As Integer = 1

            For Each para As Word.Paragraph In targetRange.Paragraphs
                Dim text As String = para.Range.Text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10))
                If String.IsNullOrWhiteSpace(text) Then Continue For

                ' Skip table cells if not processing tables
                If Not settings.ProcessTables Then
                    Try
                        If para.Range.Cells.Count > 0 Then Continue For
                    Catch
                    End Try
                End If

                ' Build formatting hints
                Dim hints As New StringBuilder()

                ' List formatting hint - detect actual bullet vs numbering
                Dim hasList As Boolean = False
                Try
                    Dim lf As Word.ListFormat = para.Range.ListFormat
                    If lf.ListType <> WdListType.wdListNoNumbering Then
                        hasList = True

                        ' Determine if this is actually a bullet list by checking the list string
                        Dim listString As String = lf.ListString
                        Dim isBullet As Boolean = False

                        If lf.ListType = WdListType.wdListBullet Then
                            isBullet = True
                        ElseIf Not String.IsNullOrEmpty(listString) Then
                            Dim trimmed As String = listString.Trim()
                            If trimmed.Length > 0 Then
                                Dim firstChar As Char = trimmed.Chars(0)
                                If Not Char.IsLetterOrDigit(firstChar) Then
                                    isBullet = True
                                End If
                            End If
                        End If

                        Dim listTypeHint As String = If(isBullet, "Bullet", lf.ListType.ToString().Replace("wdList", ""))
                        hints.Append($"[LIST:{listTypeHint}/L{lf.ListLevelNumber}] ")
                    End If
                Catch
                End Try

                ' Indentation hint
                Try
                    If para.LeftIndent > 0 OrElse para.FirstLineIndent <> 0 Then
                        hints.Append($"[INDENT:{Math.Round(para.LeftIndent, 0)}/{Math.Round(para.FirstLineIndent, 0)}] ")
                    End If
                Catch
                End Try

                paragraphTexts.AppendLine($"[{idx}] {hints}{text}")
                paragraphList.Add(para)
                paragraphHasList.Add(hasList)
                idx += 1
            Next

            If paragraphList.Count = 0 Then
                report.AppendLine("No paragraphs to process.")
                Return report.ToString()
            End If

            ' Build prompt - add numbering hint only if LLM-assisted mode is enabled
            Dim systemPrompt As String = SP_ApplyDocStyleFast
            If settings.ListNumberingReset = 2 Then
                systemPrompt &= SP_ApplyDocStyleFast_NumberingHint
            End If

            Dim userPrompt As New StringBuilder()
            userPrompt.AppendLine("<STYLETEMPLATE>")
            userPrompt.AppendLine(templateJson)
            userPrompt.AppendLine("</STYLETEMPLATE>")
            userPrompt.AppendLine()
            userPrompt.AppendLine("<DOCUMENT>")
            userPrompt.AppendLine(paragraphTexts.ToString())
            userPrompt.AppendLine("</DOCUMENT>")

            If Not String.IsNullOrWhiteSpace(settings.DocumentContext) Then
                userPrompt.AppendLine()
                userPrompt.AppendLine("<CONTEXT>")
                userPrompt.AppendLine(settings.DocumentContext)
                userPrompt.AppendLine("</CONTEXT>")
            End If

            ' Call LLM
            Dim response As String = Await LLM(systemPrompt, userPrompt.ToString(), "", "", 0, useSecondAPI)

            ' Parse response
            Dim mappingArray As JArray
            Try
                response = ExtractJsonFromResponse(response)
                mappingArray = JArray.Parse(response)
            Catch ex As Exception
                report.AppendLine($"Error parsing LLM response: {ex.Message}")
                Return report.ToString()
            End Try

            ' Apply styles
            Dim appliedCount As Integer = 0
            Dim skippedCount As Integer = 0

            ShowProgressBarInSeparateThread($"{AN} - Applying Styles", "Applying styles...")
            ProgressBarModule.CancelOperation = False
            GlobalProgressMax = mappingArray.Count
            GlobalProgressValue = 0

            ' Track paragraphs that need numbering restart (for post-processing)
            Dim numberingRestartParas As New List(Of Word.Paragraph)()

            ' TEMPLATE baseline by paragraph index (0-based index into paragraphList)
            Dim templateBaselineByIndex As New Dictionary(Of Integer, JObject)()

            Try
                For Each mapping As JObject In mappingArray
                    If ProgressBarModule.CancelOperation Then
                        report.AppendLine("Operation cancelled by user.")
                        Exit For
                    End If

                    Dim paraIdx As Integer = CInt(mapping("paragraphIndex")) - 1

                    Dim userStyleNameRaw As String = If(mapping("userStyleName") IsNot Nothing, CStr(mapping("userStyleName")), "")
                    Dim userStyleName As String = NormalizeStyleKey(userStyleNameRaw)

                    Dim confidence As Integer = If(mapping("confidence") IsNot Nothing, CInt(mapping("confidence")), 100)

                    GlobalProgressValue += 1
                    GlobalProgressLabel = $"Processing paragraph {paraIdx + 1} of {paragraphList.Count}"

                    If paraIdx < 0 OrElse paraIdx >= paragraphList.Count Then Continue For

                    Dim para As Word.Paragraph = paragraphList(paraIdx)
                    Dim paraPreview As String = GetParaPreview(para, ReportParaPreviewLen)

                    ' Resolve wdStyle name from user style name
                    Dim wdStyleName As String = "Normal"
                    If userStyleToWdStyle.ContainsKey(userStyleName) Then
                        wdStyleName = userStyleToWdStyle(userStyleName)
                    ElseIf Not String.IsNullOrWhiteSpace(userStyleName) Then
                        ' Try using the user style name directly as wdStyle name
                        wdStyleName = userStyleName
                    End If

                    ' Check confidence threshold
                    If settings.UseConfidenceThreshold AndAlso confidence < settings.ConfidenceThreshold Then
                        skippedCount += 1
                        report.AppendLine($"Skipped paragraph {paraIdx + 1}: confidence {confidence}% below threshold | suggested='{userStyleNameRaw}' norm='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                        Continue For
                    End If

                    ' Preview mode
                    If settings.PreviewMode Then
                        para.Range.Select()
                        Dim preview As Integer = ShowCustomYesNoBox(
                    $"Apply user style '{userStyleNameRaw}' (normalized: '{userStyleName}') (Word style: '{wdStyleName}') to this paragraph?{vbCrLf}{vbCrLf}Confidence: {confidence}%{vbCrLf}Text: {para.Range.Text.Substring(0, Math.Min(100, para.Range.Text.Length))}...",
                    "Yes", "Skip", $"{AN} - Preview")

                        If preview = 0 Then
                            Dim continueChoice As Integer = ShowCustomYesNoBox(
                        "You closed the preview dialog without making a selection." & vbCrLf & vbCrLf &
                        "Do you want to continue applying styles without individual preview, or abort the operation?",
                        "Continue without preview", "Abort", $"{AN} - Continue?")
                            If continueChoice = 1 Then
                                settings.PreviewMode = False
                            Else
                                report.AppendLine("Operation aborted by user.")
                                Exit For
                            End If
                        ElseIf preview <> 1 Then
                            skippedCount += 1
                            report.AppendLine($"Skipped paragraph {paraIdx + 1}: user skipped in preview | suggested='{userStyleNameRaw}' norm='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                            Continue For
                        End If
                    End If

                    ' Apply style + TEMPLATE overrides; then capture template baseline.
                    Try
                        para.Style = doc.Styles(wdStyleName)

                        ' Find template user style definition (normalized key)
                        Dim userStyleDef As JObject = Nothing
                        userStyleNameToDef.TryGetValue(userStyleName, userStyleDef)

                        If userStyleDef IsNot Nothing Then
                            ApplyUserStyleFormattingFromTemplate(para, userStyleDef)
                        Else
                            Debug.WriteLine($"[DocStyle] FastMode: No userStyleDef for para {paraIdx + 1}: raw='{userStyleNameRaw}' norm='{userStyleName}' wdStyle='{wdStyleName}' preview='{paraPreview}'")
                        End If

                        ' Capture TEMPLATE baseline AFTER wdStyle + manual template overrides.
                        templateBaselineByIndex(paraIdx) = CaptureParaProps(para)

                        appliedCount += 1

                        ' Track for numbering restart post-processing (only if feature is enabled)
                        If settings.ListNumberingReset > 0 Then
                            Dim shouldRestart As Boolean = False
                            If settings.ListNumberingReset = 2 Then
                                ' LLM-assisted: check the mapping response
                                shouldRestart = If(mapping("restartNumbering") IsNot Nothing, CBool(mapping("restartNumbering")), False)
                            End If
                            ' For rule-based (mode 1), we compute in post-processing.

                            If shouldRestart Then
                                numberingRestartParas.Add(para)
                            End If
                        End If

                    Catch ex As Exception
                        report.AppendLine($"Error applying style '{wdStyleName}' to paragraph {paraIdx + 1}: {ex.Message} | suggested='{userStyleNameRaw}' norm='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                    End Try
                Next
            Finally
                ProgressBarModule.CancelOperation = True
            End Try

            ' Post-processing: Apply numbering restarts if feature is enabled
            Dim restartCount As Integer = 0
            If settings.ListNumberingReset > 0 Then
                restartCount = ApplyNumberingRestarts(doc, paragraphList, paragraphHasList, settings, numberingRestartParas, templateBaselineByIndex)
            End If

            report.AppendLine()
            report.AppendLine($"Total paragraphs: {paragraphList.Count}")
            report.AppendLine($"Styles applied: {appliedCount}")
            report.AppendLine($"Skipped: {skippedCount}")
            If settings.ListNumberingReset > 0 Then
                report.AppendLine($"Numbering restarts applied: {restartCount}")
            End If

        Catch ex As Exception
            report.AppendLine($"Error in fast mode: {ex.Message}")
        End Try

        Return report.ToString()
    End Function


    Private Async Function oldApplyStylesFastMode(doc As Word.Document, targetRange As Word.Range,
                                                templateJson As String, templateObj As JObject,
                                                settings As DocStyleSettings,
                                                useSecondAPI As Boolean) As Task(Of String)
        Dim report As New StringBuilder()
        report.AppendLine("=== Style Template Application Report (Fast Mode) ===")
        report.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        report.AppendLine()

        Try
            ' Build user style name to wdStyle name mapping
            Dim userStyleToWdStyle As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If templateObj("userStyles") IsNot Nothing Then
                For Each userStyle As JObject In CType(templateObj("userStyles"), JArray)
                    Dim userStyleName As String = If(userStyle("userStyleName") IsNot Nothing, userStyle("userStyleName").ToString(), "")
                    Dim wdStyleName As String = If(userStyle("wdStyleName") IsNot Nothing, userStyle("wdStyleName").ToString(), "Normal")
                    If Not String.IsNullOrWhiteSpace(userStyleName) Then
                        userStyleToWdStyle(userStyleName) = wdStyleName
                    End If
                Next
            End If

            ' Build numbered paragraph list with formatting hints
            Dim paragraphTexts As New StringBuilder()
            Dim paragraphList As New List(Of Word.Paragraph)()
            Dim paragraphHasList As New List(Of Boolean)()
            Dim idx As Integer = 1

            For Each para As Word.Paragraph In targetRange.Paragraphs
                Dim text As String = para.Range.Text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10))
                If String.IsNullOrWhiteSpace(text) Then Continue For

                ' Skip table cells if not processing tables
                If Not settings.ProcessTables Then
                    Try
                        If para.Range.Cells.Count > 0 Then Continue For
                    Catch
                    End Try
                End If

                ' Build formatting hints
                Dim hints As New StringBuilder()

                ' List formatting hint - detect actual bullet vs numbering
                Dim hasList As Boolean = False
                Try
                    Dim lf As Word.ListFormat = para.Range.ListFormat
                    If lf.ListType <> WdListType.wdListNoNumbering Then
                        hasList = True

                        ' Determine if this is actually a bullet list by checking the list string
                        Dim listString As String = lf.ListString
                        Dim isBullet As Boolean = False

                        If lf.ListType = WdListType.wdListBullet Then
                            isBullet = True
                        ElseIf Not String.IsNullOrEmpty(listString) Then
                            Dim trimmed As String = listString.Trim()
                            If trimmed.Length > 0 Then
                                Dim firstChar As Char = trimmed.Chars(0)
                                If Not Char.IsLetterOrDigit(firstChar) Then
                                    isBullet = True
                                End If
                            End If
                        End If

                        Dim listTypeHint As String = If(isBullet, "Bullet", lf.ListType.ToString().Replace("wdList", ""))
                        hints.Append($"[LIST:{listTypeHint}/L{lf.ListLevelNumber}] ")
                    End If
                Catch
                End Try

                ' Indentation hint
                Try
                    If para.LeftIndent > 0 OrElse para.FirstLineIndent <> 0 Then
                        hints.Append($"[INDENT:{Math.Round(para.LeftIndent, 0)}/{Math.Round(para.FirstLineIndent, 0)}] ")
                    End If
                Catch
                End Try

                paragraphTexts.AppendLine($"[{idx}] {hints}{text}")
                paragraphList.Add(para)
                paragraphHasList.Add(hasList)
                idx += 1
            Next

            If paragraphList.Count = 0 Then
                report.AppendLine("No paragraphs to process.")
                Return report.ToString()
            End If

            ' Build prompt - add numbering hint only if LLM-assisted mode is enabled
            Dim systemPrompt As String = SP_ApplyDocStyleFast
            If settings.ListNumberingReset = 2 Then
                systemPrompt &= SP_ApplyDocStyleFast_NumberingHint
            End If

            Dim userPrompt As New StringBuilder()
            userPrompt.AppendLine("<STYLETEMPLATE>")
            userPrompt.AppendLine(templateJson)
            userPrompt.AppendLine("</STYLETEMPLATE>")
            userPrompt.AppendLine()
            userPrompt.AppendLine("<DOCUMENT>")
            userPrompt.AppendLine(paragraphTexts.ToString())
            userPrompt.AppendLine("</DOCUMENT>")

            If Not String.IsNullOrWhiteSpace(settings.DocumentContext) Then
                userPrompt.AppendLine()
                userPrompt.AppendLine("<CONTEXT>")
                userPrompt.AppendLine(settings.DocumentContext)
                userPrompt.AppendLine("</CONTEXT>")
            End If

            ' Call LLM
            Dim response As String = Await LLM(systemPrompt, userPrompt.ToString(), "", "", 0, useSecondAPI)

            ' Parse response
            Dim mappingArray As JArray
            Try
                response = ExtractJsonFromResponse(response)
                mappingArray = JArray.Parse(response)
            Catch ex As Exception
                report.AppendLine($"Error parsing LLM response: {ex.Message}")
                Return report.ToString()
            End Try

            ' Apply styles
            Dim appliedCount As Integer = 0
            Dim skippedCount As Integer = 0

            ShowProgressBarInSeparateThread($"{AN} - Applying Styles", "Applying styles...")
            ProgressBarModule.CancelOperation = False
            GlobalProgressMax = mappingArray.Count
            GlobalProgressValue = 0

            ' Track paragraphs that need numbering restart (for post-processing)
            Dim numberingRestartParas As New List(Of Word.Paragraph)()

            ' TEMPLATE baseline by paragraph index (0-based index into paragraphList)
            Dim templateBaselineByIndex As New Dictionary(Of Integer, JObject)()

            Try
                For Each mapping As JObject In mappingArray
                    If ProgressBarModule.CancelOperation Then
                        report.AppendLine("Operation cancelled by user.")
                        Exit For
                    End If

                    Dim paraIdx As Integer = CInt(mapping("paragraphIndex")) - 1
                    Dim userStyleName As String = If(mapping("userStyleName") IsNot Nothing, CStr(mapping("userStyleName")), "")
                    Dim confidence As Integer = If(mapping("confidence") IsNot Nothing, CInt(mapping("confidence")), 100)

                    GlobalProgressValue += 1
                    GlobalProgressLabel = $"Processing paragraph {paraIdx + 1} of {paragraphList.Count}"

                    If paraIdx < 0 OrElse paraIdx >= paragraphList.Count Then Continue For

                    Dim para As Word.Paragraph = paragraphList(paraIdx)
                    Dim paraPreview As String = GetParaPreview(para, ReportParaPreviewLen)

                    ' Resolve wdStyle name from user style name (do this early so we can report it even when skipped)
                    Dim wdStyleName As String = "Normal"
                    If userStyleToWdStyle.ContainsKey(userStyleName) Then
                        wdStyleName = userStyleToWdStyle(userStyleName)
                    ElseIf Not String.IsNullOrWhiteSpace(userStyleName) Then
                        ' Try using the user style name directly as wdStyle name
                        wdStyleName = userStyleName
                    End If

                    ' Check confidence threshold
                    If settings.UseConfidenceThreshold AndAlso confidence < settings.ConfidenceThreshold Then
                        skippedCount += 1
                        report.AppendLine($"Skipped paragraph {paraIdx + 1}: confidence {confidence}% below threshold | suggested='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                        Continue For
                    End If

                    ' Preview mode
                    If settings.PreviewMode Then
                        para.Range.Select()
                        Dim preview As Integer = ShowCustomYesNoBox(
                        $"Apply user style '{userStyleName}' (Word style: '{wdStyleName}') to this paragraph?{vbCrLf}{vbCrLf}Confidence: {confidence}%{vbCrLf}Text: {para.Range.Text.Substring(0, Math.Min(100, para.Range.Text.Length))}...",
                        "Yes", "Skip", $"{AN} - Preview")

                        If preview = 0 Then
                            Dim continueChoice As Integer = ShowCustomYesNoBox(
                            "You closed the preview dialog without making a selection." & vbCrLf & vbCrLf &
                            "Do you want to continue applying styles without individual preview, or abort the operation?",
                            "Continue without preview", "Abort", $"{AN} - Continue?")
                            If continueChoice = 1 Then
                                settings.PreviewMode = False
                            Else
                                report.AppendLine("Operation aborted by user.")
                                Exit For
                            End If
                        ElseIf preview <> 1 Then
                            skippedCount += 1
                            report.AppendLine($"Skipped paragraph {paraIdx + 1}: user skipped in preview | suggested='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                            Continue For
                        End If
                    End If

                    ' Apply style + TEMPLATE overrides; then capture template baseline.
                    Try
                        para.Style = doc.Styles(wdStyleName)

                        ' Find template user style definition
                        Dim userStyleDef As JObject = Nothing
                        If templateObj("userStyles") IsNot Nothing Then
                            For Each us As JObject In CType(templateObj("userStyles"), JArray)
                                If us("userStyleName") IsNot Nothing AndAlso
                               String.Equals(us("userStyleName").ToString(), userStyleName, StringComparison.OrdinalIgnoreCase) Then
                                    userStyleDef = us
                                    Exit For
                                End If
                            Next
                        End If

                        If userStyleDef IsNot Nothing Then
                            ApplyUserStyleFormattingFromTemplate(para, userStyleDef)
                        End If

                        ' Capture TEMPLATE baseline AFTER wdStyle + manual template overrides.
                        templateBaselineByIndex(paraIdx) = CaptureParaProps(para)

                        appliedCount += 1

                        ' Track for numbering restart post-processing (only if feature is enabled)
                        If settings.ListNumberingReset > 0 Then
                            Dim shouldRestart As Boolean = False
                            If settings.ListNumberingReset = 2 Then
                                ' LLM-assisted: check the mapping response
                                shouldRestart = If(mapping("restartNumbering") IsNot Nothing, CBool(mapping("restartNumbering")), False)
                            End If
                            ' For rule-based (mode 1), we compute in post-processing.

                            If shouldRestart Then
                                numberingRestartParas.Add(para)
                            End If
                        End If

                    Catch ex As Exception
                        report.AppendLine($"Error applying style '{wdStyleName}' to paragraph {paraIdx + 1}: {ex.Message} | suggested='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                    End Try
                Next
            Finally
                ProgressBarModule.CancelOperation = True
            End Try

            ' Post-processing: Apply numbering restarts if feature is enabled
            Dim restartCount As Integer = 0
            If settings.ListNumberingReset > 0 Then
                restartCount = ApplyNumberingRestarts(doc, paragraphList, paragraphHasList, settings, numberingRestartParas, templateBaselineByIndex)
            End If

            report.AppendLine()
            report.AppendLine($"Total paragraphs: {paragraphList.Count}")
            report.AppendLine($"Styles applied: {appliedCount}")
            report.AppendLine($"Skipped: {skippedCount}")
            If settings.ListNumberingReset > 0 Then
                report.AppendLine($"Numbering restarts applied: {restartCount}")
            End If

        Catch ex As Exception
            report.AppendLine($"Error in fast mode: {ex.Message}")
        End Try

        Return report.ToString()
    End Function


    ' CHANGED: More robust numbered detection.
    ' Word often gives empty/odd ListString for valid numbering (esp. outline templates).
    Private Function IsNumberedListParagraph(p As Word.Paragraph) As Boolean
        Try
            Dim lf As Word.ListFormat = p.Range.ListFormat
            If lf Is Nothing Then Return False

            Dim t As WdListType = lf.ListType
            If t = WdListType.wdListNoNumbering Then Return False

            ' Explicit bullets => not numbered
            If t = WdListType.wdListBullet OrElse t = WdListType.wdListPictureBullet Then
                Return False
            End If

            ' Treat all other list types as numbered.
            ' (wdListOutlineNumbering / wdListMixedNumbering etc.)
            Return True
        Catch
            Return False
        End Try
    End Function



    ' Returns a short preview for debug output (no paragraph mark/newlines, maxLen chars).
    Private Function GetParaPreview(p As Word.Paragraph, Optional maxLen As Integer = 25) As String
        Try
            Dim t As String = p.Range.Text
            t = t.Replace(vbCr, " ").Replace(vbLf, " ").Replace(ChrW(13), " ").Replace(ChrW(10), " ")
            t = t.Trim()

            If t.Length <= maxLen Then Return t
            Return t.Substring(0, maxLen) & "..."
        Catch
            Return ""
        End Try
    End Function


    ' Snapshot paragraph-level properties that are commonly disturbed by list operations.
    ' Intentionally does NOT include "style" because style must remain whatever was applied.
    Private Function CaptureParaProps(p As Word.Paragraph) As JObject
        Dim o As New JObject()

        Try : o("alignment") = p.Alignment.ToString() : Catch : End Try
        Try : o("leftIndent") = p.LeftIndent : Catch : End Try
        Try : o("rightIndent") = p.RightIndent : Catch : End Try
        Try : o("firstLineIndent") = p.FirstLineIndent : Catch : End Try
        Try : o("spaceBefore") = p.SpaceBefore : Catch : End Try
        Try : o("spaceAfter") = p.SpaceAfter : Catch : End Try
        Try : o("lineSpacing") = p.LineSpacing : Catch : End Try
        Try : o("lineSpacingRule") = p.LineSpacingRule.ToString() : Catch : End Try
        Try : o("keepTogether") = p.KeepTogether : Catch : End Try
        Try : o("keepWithNext") = p.KeepWithNext : Catch : End Try
        Try : o("pageBreakBefore") = p.PageBreakBefore : Catch : End Try
        Try : o("widowControl") = p.WidowControl : Catch : End Try
        Try : o("outlineLevel") = p.OutlineLevel.ToString() : Catch : End Try

        ' Tab stops are often clobbered by list operations; preserve if present.
        Try
            Dim tabs As New JArray()
            For Each ts As Word.TabStop In p.TabStops
                Dim t As New JObject()
                t("pos") = Math.Round(ts.Position, 2)
                t("align") = ts.Alignment.ToString()
                t("leader") = ts.Leader.ToString()
                tabs.Add(t)
            Next
            o("tabStops") = tabs
        Catch
        End Try

        Return o
    End Function

    Private Sub RestoreParaProps(p As Word.Paragraph, props As JObject)
        If props Is Nothing Then Exit Sub
        Try
            If props("alignment") IsNot Nothing Then p.Alignment = ParseAlignment(CStr(props("alignment")))
            If props("leftIndent") IsNot Nothing Then p.LeftIndent = CSng(props("leftIndent"))
            If props("rightIndent") IsNot Nothing Then p.RightIndent = CSng(props("rightIndent"))
            If props("firstLineIndent") IsNot Nothing Then p.FirstLineIndent = CSng(props("firstLineIndent"))
            If props("spaceBefore") IsNot Nothing Then p.SpaceBefore = CSng(props("spaceBefore"))
            If props("spaceAfter") IsNot Nothing Then p.SpaceAfter = CSng(props("spaceAfter"))
            If props("lineSpacing") IsNot Nothing Then p.LineSpacing = CSng(props("lineSpacing"))

            If props("lineSpacingRule") IsNot Nothing Then
                p.LineSpacingRule = ParseLineSpacingRule(CStr(props("lineSpacingRule")))
            End If

            If props("keepTogether") IsNot Nothing Then p.KeepTogether = CInt(props("keepTogether"))
            If props("keepWithNext") IsNot Nothing Then p.KeepWithNext = CInt(props("keepWithNext"))
            If props("pageBreakBefore") IsNot Nothing Then p.PageBreakBefore = CInt(props("pageBreakBefore"))
            If props("widowControl") IsNot Nothing Then p.WidowControl = CInt(props("widowControl"))
            If props("outlineLevel") IsNot Nothing Then p.OutlineLevel = ParseOutlineLevel(CStr(props("outlineLevel")))

            ' Restore tab stops
            If props("tabStops") IsNot Nothing AndAlso TypeOf props("tabStops") Is JArray Then
                Try
                    p.TabStops.ClearAll()
                    For Each tabDef As JObject In CType(props("tabStops"), JArray)
                        Dim pos As Single = CSng(tabDef("pos"))
                        Dim al As WdTabAlignment = ParseTabAlignment(CStr(tabDef("align")))
                        Dim ld As WdTabLeader = ParseTabLeader(CStr(tabDef("leader")))
                        p.TabStops.Add(pos, al, ld)
                    Next
                Catch
                End Try
            End If
        Catch
        End Try
    End Sub

    ' Restarts numbering for a single paragraph by re-applying its current ListTemplate
    ' with ContinuePreviousList:=False. Preserves paragraph properties (except style) to avoid list ops
    ' destroying the style/formatting.
    Private Sub RestartNumberingForParagraph(p As Word.Paragraph,
                                        Optional logicalIndex1Based As Integer = -1,
                                        Optional restoreProps As JObject = Nothing)
        Try
            Dim preview As String = GetParaPreview(p, 25)

            Dim beforeStyle As String = ""
            Try : beforeStyle = p.Style.NameLocal : Catch : End Try

            ' CHANGED: if caller provided a baseline (e.g., template baseline), use it.
            Dim propsToRestore As JObject = If(restoreProps, CaptureParaProps(p))

            Dim lf As Word.ListFormat = Nothing
            Try : lf = p.Range.ListFormat : Catch : lf = Nothing : End Try

            If lf Is Nothing OrElse lf.ListType = WdListType.wdListNoNumbering Then
                Debug.WriteLine($"[DocStyle] RestartNumbering SKIP idx={logicalIndex1Based} preview='{preview}' sig='{GetListSignature(p)}'")
                Exit Sub
            End If

            Dim tpl As Word.ListTemplate = Nothing
            Try : tpl = lf.ListTemplate : Catch : tpl = Nothing : End Try
            If tpl Is Nothing Then
                Debug.WriteLine($"[DocStyle] RestartNumbering SKIP idx={logicalIndex1Based} preview='{preview}' (no ListTemplate) sig='{GetListSignature(p)}'")
                Exit Sub
            End If

            Dim lvl As Integer = 1
            Try : lvl = lf.ListLevelNumber : Catch : lvl = 1 : End Try

            Debug.WriteLine($"[DocStyle] RestartNumbering START idx={logicalIndex1Based} preview='{preview}' style='{beforeStyle}' sig='{GetListSignature(p)}'")

            Dim applyTo As WdListApplyTo = WdListApplyTo.wdListApplyToWholeList
            Try
                p.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate:=tpl,
                ContinuePreviousList:=False,
                ApplyTo:=applyTo,
                DefaultListBehavior:=WdDefaultListBehavior.wdWord10ListBehavior,
                ApplyLevel:=lvl
            )
            Catch
                p.Range.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate:=tpl,
                ContinuePreviousList:=False,
                ApplyTo:=WdListApplyTo.wdListApplyToThisPointForward,
                DefaultListBehavior:=WdDefaultListBehavior.wdWord10ListBehavior,
                ApplyLevel:=lvl
            )
            End Try

            ' Restore the intended baseline (template baseline if provided)
            RestoreParaProps(p, propsToRestore)

            Debug.WriteLine($"[DocStyle] RestartNumbering END   idx={logicalIndex1Based} preview='{preview}' styleBefore='{beforeStyle}' styleAfter='{TryGetStyleName(p)}' sigAfter='{GetListSignature(p)}'")
        Catch ex As Exception
            Debug.WriteLine($"[DocStyle] RestartNumberingForParagraph error idx={logicalIndex1Based}: {ex.Message}")
        End Try
    End Sub

    ' --- NEW: helper for debug ---
    Private Function TryGetStyleName(p As Word.Paragraph) As String
        Try
            Return p.Style.NameLocal
        Catch
            Return ""
        End Try
    End Function


    Private Function IsHeadingLikeParagraph(p As Word.Paragraph, headingOutlineLevelMax As Integer) As Boolean
        If headingOutlineLevelMax <= 0 Then Return False
        If headingOutlineLevelMax > 9 Then headingOutlineLevelMax = 9

        Try
            Dim ol As WdOutlineLevel = p.OutlineLevel

            Select Case ol
                Case WdOutlineLevel.wdOutlineLevel1 : Return (headingOutlineLevelMax >= 1)
                Case WdOutlineLevel.wdOutlineLevel2 : Return (headingOutlineLevelMax >= 2)
                Case WdOutlineLevel.wdOutlineLevel3 : Return (headingOutlineLevelMax >= 3)
                Case WdOutlineLevel.wdOutlineLevel4 : Return (headingOutlineLevelMax >= 4)
                Case WdOutlineLevel.wdOutlineLevel5 : Return (headingOutlineLevelMax >= 5)
                Case WdOutlineLevel.wdOutlineLevel6 : Return (headingOutlineLevelMax >= 6)
                Case WdOutlineLevel.wdOutlineLevel7 : Return (headingOutlineLevelMax >= 7)
                Case WdOutlineLevel.wdOutlineLevel8 : Return (headingOutlineLevelMax >= 8)
                Case WdOutlineLevel.wdOutlineLevel9 : Return (headingOutlineLevelMax >= 9)
                Case Else
                    Return False ' BodyText
            End Select
        Catch
            Return False
        End Try
    End Function


    Private Function GuessNestingLevel(p As Word.Paragraph, baseIndent As Single) As Integer
        Try
            Dim lf As Word.ListFormat = p.Range.ListFormat
            Dim lvl As Integer = 1
            Try : lvl = lf.ListLevelNumber : Catch : lvl = 1 : End Try
            If lvl > 1 Then Return lvl

            ' Manual indent nesting: compare paragraph left indent to the run's base indent
            Dim li As Single = 0
            Try : li = CSng(p.LeftIndent) : Catch : li = 0 : End Try

            Dim delta As Single = li - baseIndent
            If delta <= (IndentStepPoints * 0.5F) Then Return 1

            Dim inferred As Integer = 1 + CInt(Math.Floor(delta / IndentStepPoints))
            If inferred < 1 Then inferred = 1
            If inferred > 9 Then inferred = 9
            Return inferred
        Catch
            Return 1
        End Try
    End Function


    Private Function ApplyNumberingRestarts(doc As Word.Document,
                                       paragraphList As List(Of Word.Paragraph),
                                       paragraphHasList As List(Of Boolean),
                                       settings As DocStyleSettings,
                                       llmRestartParas As List(Of Word.Paragraph),
                                       templateBaselineByIndex As Dictionary(Of Integer, JObject)) As Integer
        Dim restartCount As Integer = 0

        Try
            Dim restartIndices As New HashSet(Of Integer)()

            If settings.ListNumberingReset = 1 Then
                Dim inBodyNumberedRun As Boolean = False
                Dim restartedMainThisRun As Boolean = False

                Dim lastSeenIndexByLevel As New Dictionary(Of Integer, Integer)()
                Dim restartedForParentKey As New HashSet(Of String)(StringComparer.Ordinal)

                Dim runBaseIndent As Single = 0
                Dim haveRunBaseIndent As Boolean = False

                For i As Integer = 0 To paragraphList.Count - 1
                    Dim p = paragraphList(i)

                    Dim isNumbered As Boolean = IsNumberedListParagraph(p)
                    Dim isBullet As Boolean = If(RuleBreakOnBullets, IsBulletListParagraph(p), False)
                    Dim isHeadingLike As Boolean = IsHeadingLikeParagraph(p, settings.RuleHeadingOutlineLevelMax)

                    If (Not isNumbered) OrElse isBullet OrElse isHeadingLike Then
                        inBodyNumberedRun = False
                        restartedMainThisRun = False
                        lastSeenIndexByLevel.Clear()
                        restartedForParentKey.Clear()
                        haveRunBaseIndent = False
                        runBaseIndent = 0
                        Continue For
                    End If

                    If Not inBodyNumberedRun Then
                        inBodyNumberedRun = True
                        restartedMainThisRun = False
                        lastSeenIndexByLevel.Clear()
                        restartedForParentKey.Clear()
                        haveRunBaseIndent = False
                        runBaseIndent = 0
                    End If

                    If Not haveRunBaseIndent Then
                        Try : runBaseIndent = CSng(p.LeftIndent) : Catch : runBaseIndent = 0 : End Try
                        haveRunBaseIndent = True
                    End If

                    Dim effLevel As Integer = GuessNestingLevel(p, runBaseIndent)

                    Dim parentIdx As Integer = -1
                    If effLevel > 1 Then
                        If lastSeenIndexByLevel.ContainsKey(effLevel - 1) Then
                            parentIdx = lastSeenIndexByLevel(effLevel - 1)
                        Else
                            For pl As Integer = effLevel - 2 To 1 Step -1
                                If lastSeenIndexByLevel.ContainsKey(pl) Then
                                    parentIdx = lastSeenIndexByLevel(pl)
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    If effLevel = 1 Then
                        If Not restartedMainThisRun Then
                            restartIndices.Add(i)
                            restartedMainThisRun = True
                            Debug.WriteLine($"[DocStyle] RuleRestart(main) idx={i + 1} effLevel=1 leftIndent={Math.Round(p.LeftIndent, 1)} preview='{GetParaPreview(p, 25)}' sig='{GetListSignature(p)}'")
                        End If
                    Else
                        Dim key As String = $"{effLevel}:{parentIdx}"
                        If parentIdx >= 0 AndAlso Not restartedForParentKey.Contains(key) Then
                            restartIndices.Add(i)
                            restartedForParentKey.Add(key)
                            Debug.WriteLine($"[DocStyle] RuleRestart(sub) idx={i + 1} effLevel={effLevel} parentIdx={parentIdx + 1} leftIndent={Math.Round(p.LeftIndent, 1)} preview='{GetParaPreview(p, 25)}' sig='{GetListSignature(p)}'")
                        ElseIf parentIdx < 0 Then
                            Dim fallbackKey As String = $"{effLevel}:-1"
                            If Not restartedForParentKey.Contains(fallbackKey) Then
                                restartIndices.Add(i)
                                restartedForParentKey.Add(fallbackKey)
                                Debug.WriteLine($"[DocStyle] RuleRestart(sub-fallback) idx={i + 1} effLevel={effLevel} leftIndent={Math.Round(p.LeftIndent, 1)} preview='{GetParaPreview(p, 25)}' sig='{GetListSignature(p)}'")
                            End If
                        End If
                    End If

                    lastSeenIndexByLevel(effLevel) = i

                    Dim deeper = lastSeenIndexByLevel.Keys.Where(Function(k) k > effLevel).ToList()
                    For Each k In deeper
                        lastSeenIndexByLevel.Remove(k)
                    Next
                Next

            ElseIf settings.ListNumberingReset = 2 Then
                For i As Integer = 0 To paragraphList.Count - 1
                    If llmRestartParas.Contains(paragraphList(i)) AndAlso IsNumberedListParagraph(paragraphList(i)) Then
                        restartIndices.Add(i)
                    End If
                Next
            End If

            For Each idx In restartIndices.OrderBy(Function(x) x)
                Dim p As Word.Paragraph = paragraphList(idx)

                Dim baseline As JObject = Nothing
                If templateBaselineByIndex IsNot Nothing Then
                    templateBaselineByIndex.TryGetValue(idx, baseline)
                End If

                Debug.WriteLine($"[DocStyle] ApplyNumberingRestarts restarting paragraph {idx + 1} preview='{GetParaPreview(p, 25)}' sig='{GetListSignature(p)}'")

                ' Restart using template baseline restore
                RestartNumberingForParagraph(p, idx + 1, baseline)
                restartCount += 1

                ' REPAIR: Word can perturb sibling paragraphs in the same list/template after a restart.
                ' Apply baseline to paragraphs that share the same ListTemplate instance as the restarted paragraph.
                Try
                    Dim restartedTplId As Integer = GetListTemplateId(p)
                    If restartedTplId <> 0 AndAlso templateBaselineByIndex IsNot Nothing Then
                        For j As Integer = 0 To paragraphList.Count - 1
                            If j = idx Then Continue For

                            Dim pj As Word.Paragraph = paragraphList(j)

                            ' Same list template instance? Then reapply the template baseline props we captured earlier.
                            If GetListTemplateId(pj) = restartedTplId Then
                                Dim bj As JObject = Nothing
                                If templateBaselineByIndex.TryGetValue(j, bj) AndAlso bj IsNot Nothing Then
                                    RestoreParaProps(pj, bj)
                                End If
                            End If
                        Next
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[DocStyle] Repair after restart failed: {ex.Message}")
                End Try
            Next

        Catch ex As Exception
            Debug.WriteLine($"Error in ApplyNumberingRestarts: {ex.Message}")
        End Try

        Return restartCount
    End Function

    Private Function IsBulletListParagraph(p As Word.Paragraph) As Boolean
        Try
            Dim lf = p.Range.ListFormat
            If lf Is Nothing Then Return False
            Return lf.ListType = WdListType.wdListBullet OrElse lf.ListType = WdListType.wdListPictureBullet
        Catch
            Return False
        End Try
    End Function

    ' Finds the previous paragraph index that has non-empty text (ignoring pure paragraph marks).
    Private Function FindPrevMeaningfulIndex(paragraphList As List(Of Word.Paragraph), startIdx As Integer) As Integer
        For j As Integer = startIdx - 1 To 0 Step -1
            Try
                Dim t = paragraphList(j).Range.Text
                t = t.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10)).Trim()
                If Not String.IsNullOrWhiteSpace(t) Then Return j
            Catch
            End Try
        Next
        Return -1
    End Function

    Private Function GetListSignature(p As Word.Paragraph) As String
        Try
            Dim lf = p.Range.ListFormat
            If lf Is Nothing OrElse lf.ListType = WdListType.wdListNoNumbering Then Return "NoList"

            Dim tplHash As String = "tpl=?"
            Try
                ' Not stable across sessions, but good enough for in-run debug
                tplHash = $"tpl#{RuntimeHelpers.GetHashCode(lf.ListTemplate)}"
            Catch
            End Try

            Dim lvl As Integer = 1
            Try : lvl = lf.ListLevelNumber : Catch : End Try

            Dim ls As String = ""
            Try : ls = lf.ListString : Catch : End Try

            Return $"{lf.ListType}/{tplHash}/L{lvl}/'{ls}'"
        Catch
            Return "ListSigError"
        End Try
    End Function

    ' --- NEW: helper to detect "has any list" quickly (bullets OR numbering) ---
    Private Function HasAnyList(p As Word.Paragraph) As Boolean
        Try
            Return p.Range.ListFormat IsNot Nothing AndAlso p.Range.ListFormat.ListType <> WdListType.wdListNoNumbering
        Catch
            Return False
        End Try
    End Function

    ' --- NEW: helper for template identity (best-effort) ---
    Private Function GetListTemplateId(p As Word.Paragraph) As Integer
        Try
            Dim tpl = p.Range.ListFormat.ListTemplate
            If tpl Is Nothing Then Return 0
            Return RuntimeHelpers.GetHashCode(tpl)
        Catch
            Return 0
        End Try
    End Function


    ''' <summary>
    ''' Restarts numbering for a range of paragraphs.
    ''' Creates a new list starting at the first paragraph and links all subsequent ones to it.
    ''' </summary>
    Private Sub RestartNumberingForRange(paragraphList As List(Of Word.Paragraph),
                                          startIdx As Integer, endIdx As Integer)
        Try
            Dim startPara As Word.Paragraph = paragraphList(startIdx)

            ' Save formatting for all paragraphs in range
            Dim savedFormatting As New List(Of (
                alignment As WdParagraphAlignment,
                leftIndent As Single,
                rightIndent As Single,
                firstLineIndent As Single,
                spaceBefore As Single,
                spaceAfter As Single,
                lineSpacing As Single,
                lineSpacingRule As WdLineSpacing,
                listLevel As Integer))()

            For i As Integer = startIdx To endIdx
                Try
                    Dim para As Word.Paragraph = paragraphList(i)
                    Dim lvl As Integer = 1
                    Try : lvl = para.Range.ListFormat.ListLevelNumber : Catch : End Try

                    savedFormatting.Add((
                        para.Alignment,
                        para.LeftIndent,
                        para.RightIndent,
                        para.FirstLineIndent,
                        para.SpaceBefore,
                        para.SpaceAfter,
                        para.LineSpacing,
                        para.LineSpacingRule,
                        lvl))
                Catch
                    savedFormatting.Add((WdParagraphAlignment.wdAlignParagraphLeft, 0, 0, 0, 0, 0, 12, WdLineSpacing.wdLineSpaceSingle, 1))
                End Try
            Next

            ' Build a range spanning all paragraphs that should be in this list
            Dim rangeStart As Integer = startPara.Range.Start
            Dim rangeEnd As Integer = paragraphList(endIdx).Range.End
            Dim combinedRange As Word.Range = startPara.Range.Document.Range(rangeStart, rangeEnd)

            ' Remove existing numbering from all paragraphs in range
            combinedRange.ListFormat.RemoveNumbers()

            ' Apply new numbering to the combined range - this creates a single new list
            combinedRange.ListFormat.ApplyNumberDefault()

            ' Restore formatting and list levels
            For i As Integer = startIdx To endIdx
                Try
                    Dim para As Word.Paragraph = paragraphList(i)
                    Dim saved = savedFormatting(i - startIdx)

                    ' Restore paragraph formatting
                    para.Alignment = saved.alignment
                    para.LeftIndent = saved.leftIndent
                    para.RightIndent = saved.rightIndent
                    para.FirstLineIndent = saved.firstLineIndent
                    para.SpaceBefore = saved.spaceBefore
                    para.SpaceAfter = saved.spaceAfter
                    para.LineSpacing = saved.lineSpacing
                    para.LineSpacingRule = saved.lineSpacingRule

                    ' Restore list level
                    Dim currentLevel As Integer = 1
                    Try : currentLevel = para.Range.ListFormat.ListLevelNumber : Catch : End Try

                    If currentLevel < saved.listLevel Then
                        For lvl As Integer = currentLevel + 1 To saved.listLevel
                            para.Range.ListFormat.ListIndent()
                        Next
                    ElseIf currentLevel > saved.listLevel Then
                        For lvl As Integer = currentLevel - 1 To saved.listLevel Step -1
                            para.Range.ListFormat.ListOutdent()
                        Next
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error restoring formatting for paragraph {i}: {ex.Message}")
                End Try
            Next

            Debug.WriteLine($"Restarted numbering for paragraphs {startIdx + 1} to {endIdx + 1}")

        Catch ex As Exception
            Debug.WriteLine($"Error in RestartNumberingForRange: {ex.Message}")
        End Try
    End Sub





    ''' <summary>
    ''' Applies styles using detailed mode (paragraph by paragraph with full formatting).
    ''' </summary>
    Private Async Function ApplyStylesDetailedMode(doc As Word.Document, targetRange As Word.Range,
                                                    templateJson As String, templateObj As JObject,
                                                    settings As DocStyleSettings,
                                                    useSecondAPI As Boolean) As Task(Of String)
        Dim report As New StringBuilder()
        report.AppendLine("=== Style Template Application Report (Detailed Mode) ===")
        report.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        report.AppendLine()

        Try
            ' Build user style name to wdStyle name mapping
            Dim userStyleToWdStyle As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If templateObj("userStyles") IsNot Nothing Then
                For Each userStyle As JObject In CType(templateObj("userStyles"), JArray)
                    Dim userStyleName As String = If(userStyle("userStyleName") IsNot Nothing, userStyle("userStyleName").ToString(), "")
                    Dim wdStyleName As String = If(userStyle("wdStyleName") IsNot Nothing, userStyle("wdStyleName").ToString(), "Normal")
                    If Not String.IsNullOrWhiteSpace(userStyleName) Then
                        userStyleToWdStyle(userStyleName) = wdStyleName
                    End If
                Next
            End If

            ' Collect paragraphs to process
            Dim paragraphList As New List(Of Word.Paragraph)()
            For Each para As Word.Paragraph In targetRange.Paragraphs
                Dim text As String = para.Range.Text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10))
                If String.IsNullOrWhiteSpace(text) Then Continue For

                If Not settings.ProcessTables Then
                    Try
                        If para.Range.Cells.Count > 0 Then Continue For
                    Catch
                    End Try
                End If

                paragraphList.Add(para)
            Next

            If paragraphList.Count = 0 Then
                report.AppendLine("No paragraphs to process.")
                Return report.ToString()
            End If

            ' Build full document text for context
            Dim fullDocText As String = targetRange.Text

            ' Process paragraphs
            Dim appliedCount As Integer = 0
            Dim skippedCount As Integer = 0

            ShowProgressBarInSeparateThread($"{AN} - Applying Styles", "Analyzing paragraphs...")
            ProgressBarModule.CancelOperation = False
            GlobalProgressMax = paragraphList.Count
            GlobalProgressValue = 0

            Try
                For i As Integer = 0 To paragraphList.Count - 1
                    If ProgressBarModule.CancelOperation Then
                        report.AppendLine("Operation cancelled by user.")
                        Exit For
                    End If

                    GlobalProgressValue = i + 1
                    GlobalProgressLabel = $"Processing paragraph {i + 1} of {paragraphList.Count}"

                    Dim para As Word.Paragraph = paragraphList(i)
                    Dim paraText As String = para.Range.Text.TrimEnd(vbCr, vbLf, ChrW(13), ChrW(10))
                    Dim paraPreview As String = GetParaPreview(para, ReportParaPreviewLen)

                    ' Get current formatting as JSON (used for LLM context)
                    Dim currentProps As JObject = ExtractParagraphFormattingForLLM(para, i + 1)

                    ' Capture original list formatting before any changes
                    Dim originalHasList As Boolean = False
                    Dim originalListType As WdListType = WdListType.wdListNoNumbering
                    Dim originalListLevel As Integer = 1
                    Try
                        originalListType = para.Range.ListFormat.ListType
                        originalHasList = (originalListType <> WdListType.wdListNoNumbering)
                        If originalHasList Then
                            originalListLevel = para.Range.ListFormat.ListLevelNumber
                        End If
                    Catch
                    End Try

                    ' Build prompt
                    Dim userPrompt As New StringBuilder()
                    userPrompt.AppendLine("<STYLETEMPLATE>")
                    userPrompt.AppendLine(templateJson)
                    userPrompt.AppendLine("</STYLETEMPLATE>")
                    userPrompt.AppendLine()
                    userPrompt.AppendLine("<CURRENTPARA>")
                    userPrompt.AppendLine(paraText)
                    userPrompt.AppendLine("</CURRENTPARA>")
                    userPrompt.AppendLine()
                    userPrompt.AppendLine("<CURRENTPROPS>")
                    userPrompt.AppendLine(currentProps.ToString(Formatting.None))
                    userPrompt.AppendLine("</CURRENTPROPS>")
                    userPrompt.AppendLine()
                    userPrompt.AppendLine("<DOCUMENT>")
                    userPrompt.AppendLine(fullDocText)
                    userPrompt.AppendLine("</DOCUMENT>")

                    If Not String.IsNullOrWhiteSpace(settings.DocumentContext) Then
                        userPrompt.AppendLine()
                        userPrompt.AppendLine("<CONTEXT>")
                        userPrompt.AppendLine(settings.DocumentContext)
                        userPrompt.AppendLine("</CONTEXT>")
                    End If

                    ' Call LLM
                    Dim response As String = Await LLM(SP_ApplyDocStyle, userPrompt.ToString(), "", "", 0, useSecondAPI)

                    ' Parse response
                    Dim formatSpec As JObject
                    Try
                        response = ExtractJsonFromResponse(response)
                        formatSpec = JObject.Parse(response)
                    Catch ex As Exception
                        report.AppendLine($"Paragraph {i + 1}: Error parsing LLM response - {ex.Message} | text='{paraPreview}'")
                        Continue For
                    End Try

                    Dim userStyleName As String = If(formatSpec("userStyleName") IsNot Nothing, CStr(formatSpec("userStyleName")), "")
                    Dim confidence As Integer = If(formatSpec("confidence") IsNot Nothing, CInt(formatSpec("confidence")), 100)
                    Dim preserveListFormatting As Boolean = If(formatSpec("preserveListFormatting") IsNot Nothing, CBool(formatSpec("preserveListFormatting")), False)
                    Dim reasoning As String = If(formatSpec("reasoning") IsNot Nothing, CStr(formatSpec("reasoning")), "")

                    ' Resolve wdStyle name from user style name (do this early so we can report it even when skipped)
                    Dim wdStyleName As String = "Normal"
                    If userStyleToWdStyle.ContainsKey(userStyleName) Then
                        wdStyleName = userStyleToWdStyle(userStyleName)
                    ElseIf Not String.IsNullOrWhiteSpace(userStyleName) Then
                        wdStyleName = userStyleName
                    End If

                    ' Check confidence threshold
                    If settings.UseConfidenceThreshold AndAlso confidence < settings.ConfidenceThreshold Then
                        skippedCount += 1
                        report.AppendLine($"Paragraph {i + 1}: Skipped (confidence {confidence}% below threshold) | suggested='{userStyleName}' wdStyle='{wdStyleName}' | reasoning='{reasoning}' | text='{paraPreview}'")
                        Continue For
                    End If

                    ' Preview mode
                    If settings.PreviewMode Then
                        para.Range.Select()
                        Dim preview As Integer = ShowCustomYesNoBox(
                            $"Apply user style '{userStyleName}' (Word style: '{wdStyleName}') to this paragraph?{vbCrLf}{vbCrLf}Confidence: {confidence}%{vbCrLf}Reasoning: {reasoning}{vbCrLf}Preserve list: {preserveListFormatting}{vbCrLf}{vbCrLf}Text: {paraText.Substring(0, Math.Min(100, paraText.Length))}...",
                            "Yes", "Skip", $"{AN} - Preview")

                        If preview = 0 Then
                            ' User closed dialog without choosing - ask what to do
                            Dim continueChoice As Integer = ShowCustomYesNoBox(
                                "You closed the preview dialog without making a selection." & vbCrLf & vbCrLf &
                                "Do you want to continue applying styles without individual preview, or abort the operation?",
                                "Continue without preview", "Abort", $"{AN} - Continue?")
                            If continueChoice = 1 Then
                                ' Continue without preview - disable preview mode for remaining paragraphs
                                settings.PreviewMode = False
                                ' Fall through to apply style to current paragraph
                            Else
                                ' Abort the operation (user clicked Abort or closed dialog)
                                report.AppendLine("Operation aborted by user.")
                                Exit For
                            End If
                        ElseIf preview <> 1 Then
                            ' User clicked Skip
                            skippedCount += 1
                            report.AppendLine($"Paragraph {i + 1}: Skipped (user skipped in preview) | suggested='{userStyleName}' wdStyle='{wdStyleName}' | reasoning='{reasoning}' | text='{paraPreview}'")
                            Continue For
                        End If
                    End If

                    ' Apply formatting
                    Try
                        ApplyFormattingToParagraphEx(doc, para, wdStyleName, formatSpec("formatting"),
                                                      preserveListFormatting OrElse originalHasList,
                                                      originalListType, originalListLevel)
                        appliedCount += 1
                    Catch ex As Exception
                        report.AppendLine($"Paragraph {i + 1}: Error applying formatting - {ex.Message} | suggested='{userStyleName}' wdStyle='{wdStyleName}' | text='{paraPreview}'")
                    End Try
                Next
            Finally
                ProgressBarModule.CancelOperation = True
            End Try

            report.AppendLine()
            report.AppendLine($"Total paragraphs: {paragraphList.Count}")
            report.AppendLine($"Formatting applied: {appliedCount}")
            report.AppendLine($"Skipped: {skippedCount}")

        Catch ex As Exception
            report.AppendLine($"Error in detailed mode: {ex.Message}")
        End Try

        Return report.ToString()
    End Function


    ''' <summary>
    ''' Creates a minimal style template for LLM consumption.
    ''' Only includes userStyleName and whenToApply - no formatting details.
    ''' </summary>
    Private Function CreateMinimalTemplateForLLM(templateObj As JObject) As String
        Dim minimal As New JObject()
        minimal("templateName") = templateObj("templateName")

        Dim minimalStyles As New JArray()
        If templateObj("userStyles") IsNot Nothing Then
            For Each userStyle As JObject In CType(templateObj("userStyles"), JArray)
                Dim minStyle As New JObject()
                minStyle("userStyleName") = userStyle("userStyleName")
                minStyle("whenToApply") = userStyle("whenToApply")
                ' Optionally include hasList hint so LLM knows if style expects list
                If userStyle("listFormatting") IsNot Nothing AndAlso
                   userStyle("listFormatting")("hasList") IsNot Nothing Then
                    minStyle("hasList") = userStyle("listFormatting")("hasList")
                End If
                minimalStyles.Add(minStyle)
            Next
        End If
        minimal("userStyles") = minimalStyles

        Return minimal.ToString(Formatting.None)
    End Function

    ''' <summary>
    ''' Extracts paragraph formatting in a compact format optimized for LLM consumption during Step 2.
    ''' Focuses on properties that help identify paragraph purpose: lists, indentation, outline level.
    ''' </summary>
    Private Function ExtractParagraphFormattingForLLM(para As Word.Paragraph, index As Integer) As JObject
        Dim result As New JObject()
        result("index") = index

        ' List formatting - critical for understanding document structure
        Try
            Dim listFormat As Word.ListFormat = para.Range.ListFormat
            If listFormat.ListType <> WdListType.wdListNoNumbering Then
                Dim lf As New JObject()
                lf("hasList") = True
                lf("listType") = listFormat.ListType.ToString().Replace("wdList", "")
                lf("listLevel") = listFormat.ListLevelNumber
                lf("listString") = listFormat.ListString
                result("listFormatting") = lf
            Else
                result("listFormatting") = New JObject From {{"hasList", False}}
            End If
        Catch
            result("listFormatting") = New JObject From {{"hasList", False}}
        End Try

        ' Key paragraph formatting for structure identification
        Try
            Dim pf As New JObject()
            pf("leftIndent") = Math.Round(para.LeftIndent, 1)
            pf("firstLineIndent") = Math.Round(para.FirstLineIndent, 1)
            pf("outlineLevel") = para.OutlineLevel.ToString().Replace("wdOutlineLevel", "")
            pf("alignment") = para.Alignment.ToString().Replace("wdAlignParagraph", "")
            result("paragraphFormatting") = pf
        Catch
        End Try

        ' Basic font info for heading detection
        Try
            Dim font As Word.Font = para.Range.Font
            Dim ff As New JObject()
            If font.Bold = -1 Then ff("bold") = True
            If font.Italic = -1 Then ff("italic") = True
            If font.Size <> CSng(WdConstants.wdUndefined) AndAlso font.Size > 12 Then
                ff("fontSize") = font.Size
            End If
            If ff.Count > 0 Then
                result("fontFormatting") = ff
            End If
        Catch
        End Try

        ' Style info
        Try
            result("currentWdStyle") = para.Style.NameLocal
        Catch
        End Try

        ' Table context
        Try
            If para.Range.Cells.Count > 0 Then
                result("isInTable") = True
            End If
        Catch
        End Try

        Return result
    End Function


    ''' <summary>
    ''' Applies all formatting to a newly created style.
    ''' </summary>
    Private Sub ApplyAllFormattingToNewStyle(doc As Word.Document, style As Word.Style, styleDef As JObject)
        ' Apply paragraph formatting
        If styleDef("paragraphFormat") IsNot Nothing Then
            ApplyParagraphFormatToStyle(style, styleDef("paragraphFormat"))
        End If

        ' Apply tab stops
        If styleDef("tabStops") IsNot Nothing Then
            ApplyTabStopsToStyle(style, CType(styleDef("tabStops"), JArray))
        End If

        ' Apply font formatting
        If styleDef("fontFormat") IsNot Nothing Then
            ApplyFontFormatToStyle(style, styleDef("fontFormat"))
        End If

        ' Apply list formatting
        If styleDef("listFormat") IsNot Nothing Then
            ApplyListFormatToStyle(doc, style, styleDef("listFormat"))
        End If
    End Sub

    ''' <summary>
    ''' Compares and applies paragraph format only if different. Returns True if changes were made.
    ''' </summary>
    Private Function ApplyParagraphFormatIfDifferent(style As Word.Style, pf As JToken) As Boolean
        Dim changesMade As Boolean = False
        Dim paraFormat As Word.ParagraphFormat = style.ParagraphFormat

        Try
            If pf("alignment") IsNot Nothing Then
                Dim desired As WdParagraphAlignment = ParseAlignment(CStr(pf("alignment")))
                If paraFormat.Alignment <> desired Then
                    paraFormat.Alignment = desired
                    changesMade = True
                End If
            End If

            If pf("leftIndent") IsNot Nothing Then
                Dim desired As Single = CSng(pf("leftIndent"))
                If Math.Abs(paraFormat.LeftIndent - desired) > 0.1 Then
                    paraFormat.LeftIndent = desired
                    changesMade = True
                End If
            End If

            If pf("rightIndent") IsNot Nothing Then
                Dim desired As Single = CSng(pf("rightIndent"))
                If Math.Abs(paraFormat.RightIndent - desired) > 0.1 Then
                    paraFormat.RightIndent = desired
                    changesMade = True
                End If
            End If

            If pf("firstLineIndent") IsNot Nothing Then
                Dim desired As Single = CSng(pf("firstLineIndent"))
                If Math.Abs(paraFormat.FirstLineIndent - desired) > 0.1 Then
                    paraFormat.FirstLineIndent = desired
                    changesMade = True
                End If
            End If

            If pf("spaceBefore") IsNot Nothing Then
                Dim desired As Single = CSng(pf("spaceBefore"))
                If Math.Abs(paraFormat.SpaceBefore - desired) > 0.1 Then
                    paraFormat.SpaceBefore = desired
                    changesMade = True
                End If
            End If

            If pf("spaceAfter") IsNot Nothing Then
                Dim desired As Single = CSng(pf("spaceAfter"))
                If Math.Abs(paraFormat.SpaceAfter - desired) > 0.1 Then
                    paraFormat.SpaceAfter = desired
                    changesMade = True
                End If
            End If

            If pf("lineSpacing") IsNot Nothing Then
                Dim desired As Single = CSng(pf("lineSpacing"))
                If Math.Abs(paraFormat.LineSpacing - desired) > 0.1 Then
                    paraFormat.LineSpacing = desired
                    changesMade = True
                End If
            End If

            If pf("lineSpacingRule") IsNot Nothing Then
                Dim desired As WdLineSpacing = ParseLineSpacingRule(CStr(pf("lineSpacingRule")))
                If paraFormat.LineSpacingRule <> desired Then
                    paraFormat.LineSpacingRule = desired
                    changesMade = True
                End If
            End If

            If pf("keepTogether") IsNot Nothing Then
                Dim desired As Integer = CInt(pf("keepTogether"))
                If paraFormat.KeepTogether <> desired Then
                    paraFormat.KeepTogether = desired
                    changesMade = True
                End If
            End If

            If pf("keepWithNext") IsNot Nothing Then
                Dim desired As Integer = CInt(pf("keepWithNext"))
                If paraFormat.KeepWithNext <> desired Then
                    paraFormat.KeepWithNext = desired
                    changesMade = True
                End If
            End If

            If pf("pageBreakBefore") IsNot Nothing Then
                Dim desired As Integer = CInt(pf("pageBreakBefore"))
                If paraFormat.PageBreakBefore <> desired Then
                    paraFormat.PageBreakBefore = desired
                    changesMade = True
                End If
            End If

            If pf("widowControl") IsNot Nothing Then
                Dim desired As Integer = CInt(pf("widowControl"))
                If paraFormat.WidowControl <> desired Then
                    paraFormat.WidowControl = desired
                    changesMade = True
                End If
            End If

            If pf("outlineLevel") IsNot Nothing Then
                Dim desired As WdOutlineLevel = ParseOutlineLevel(CStr(pf("outlineLevel")))
                If paraFormat.OutlineLevel <> desired Then
                    paraFormat.OutlineLevel = desired
                    changesMade = True
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"Error comparing/applying paragraph format: {ex.Message}")
        End Try

        Return changesMade
    End Function

    ''' <summary>
    ''' Applies paragraph format to style without comparison (for new styles).
    ''' </summary>
    Private Sub ApplyParagraphFormatToStyle(style As Word.Style, pf As JToken)
        Dim paraFormat As Word.ParagraphFormat = style.ParagraphFormat

        Try
            If pf("alignment") IsNot Nothing Then paraFormat.Alignment = ParseAlignment(CStr(pf("alignment")))
            If pf("leftIndent") IsNot Nothing Then paraFormat.LeftIndent = CSng(pf("leftIndent"))
            If pf("rightIndent") IsNot Nothing Then paraFormat.RightIndent = CSng(pf("rightIndent"))
            If pf("firstLineIndent") IsNot Nothing Then paraFormat.FirstLineIndent = CSng(pf("firstLineIndent"))
            If pf("spaceBefore") IsNot Nothing Then paraFormat.SpaceBefore = CSng(pf("spaceBefore"))
            If pf("spaceAfter") IsNot Nothing Then paraFormat.SpaceAfter = CSng(pf("spaceAfter"))
            If pf("lineSpacing") IsNot Nothing Then paraFormat.LineSpacing = CSng(pf("lineSpacing"))
            If pf("lineSpacingRule") IsNot Nothing Then paraFormat.LineSpacingRule = ParseLineSpacingRule(CStr(pf("lineSpacingRule")))
            If pf("keepTogether") IsNot Nothing Then paraFormat.KeepTogether = CInt(pf("keepTogether"))
            If pf("keepWithNext") IsNot Nothing Then paraFormat.KeepWithNext = CInt(pf("keepWithNext"))
            If pf("pageBreakBefore") IsNot Nothing Then paraFormat.PageBreakBefore = CInt(pf("pageBreakBefore"))
            If pf("widowControl") IsNot Nothing Then paraFormat.WidowControl = CInt(pf("widowControl"))
            If pf("outlineLevel") IsNot Nothing Then paraFormat.OutlineLevel = ParseOutlineLevel(CStr(pf("outlineLevel")))
        Catch ex As Exception
            Debug.WriteLine($"Error applying paragraph format: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Compares and applies tab stops only if different. Returns True if changes were made.
    ''' </summary>
    Private Function ApplyTabStopsIfDifferent(style As Word.Style, tabsDef As JArray) As Boolean
        Try
            ' Build current tab stops signature
            Dim currentTabs As New List(Of String)()
            For Each tabStop As Word.TabStop In style.ParagraphFormat.TabStops
                currentTabs.Add($"{Math.Round(tabStop.Position, 1)}:{tabStop.Alignment}:{tabStop.Leader}")
            Next

            ' Build desired tab stops signature
            Dim desiredTabs As New List(Of String)()
            For Each tabDef As JObject In tabsDef
                Dim pos As Single = CSng(tabDef("position"))
                Dim align As String = If(tabDef("alignment") IsNot Nothing, CStr(tabDef("alignment")), "wdAlignTabLeft")
                Dim leader As String = If(tabDef("leader") IsNot Nothing, CStr(tabDef("leader")), "wdTabLeaderSpaces")
                desiredTabs.Add($"{Math.Round(pos, 1)}:{ParseTabAlignment(align)}:{ParseTabLeader(leader)}")
            Next

            ' Compare
            If currentTabs.Count = desiredTabs.Count AndAlso
               currentTabs.SequenceEqual(desiredTabs) Then
                Return False ' No changes needed
            End If

            ' Apply changes
            ApplyTabStopsToStyle(style, tabsDef)
            Return True

        Catch ex As Exception
            Debug.WriteLine($"Error comparing tab stops: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Applies tab stops to style without comparison.
    ''' </summary>
    Private Sub ApplyTabStopsToStyle(style As Word.Style, tabsDef As JArray)
        Try
            style.ParagraphFormat.TabStops.ClearAll()
            For Each tabDef As JObject In tabsDef
                Dim position As Single = CSng(tabDef("position"))
                Dim alignment As WdTabAlignment = ParseTabAlignment(If(tabDef("alignment") IsNot Nothing, CStr(tabDef("alignment")), "wdAlignTabLeft"))
                Dim leader As WdTabLeader = ParseTabLeader(If(tabDef("leader") IsNot Nothing, CStr(tabDef("leader")), "wdTabLeaderSpaces"))
                style.ParagraphFormat.TabStops.Add(position, alignment, leader)
            Next
        Catch ex As Exception
            Debug.WriteLine($"Error applying tab stops: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Compares and applies font format only if different. Returns True if changes were made.
    ''' </summary>
    Private Function ApplyFontFormatIfDifferent(style As Word.Style, ff As JToken) As Boolean
        Dim changesMade As Boolean = False
        Dim font As Word.Font = style.Font

        Try
            If ff("name") IsNot Nothing Then
                Dim desired As String = CStr(ff("name"))
                If font.Name <> desired Then
                    font.Name = desired
                    changesMade = True
                End If
            End If

            If ff("size") IsNot Nothing Then
                Dim desired As Single = CSng(ff("size"))
                If Math.Abs(font.Size - desired) > 0.1 Then
                    font.Size = desired
                    changesMade = True
                End If
            End If

            If ff("bold") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("bold")), -1, 0)
                If font.Bold <> desired Then
                    font.Bold = desired
                    changesMade = True
                End If
            End If

            If ff("italic") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("italic")), -1, 0)
                If font.Italic <> desired Then
                    font.Italic = desired
                    changesMade = True
                End If
            End If

            If ff("underline") IsNot Nothing Then
                Dim desired As WdUnderline = ParseUnderline(CStr(ff("underline")))
                If font.Underline <> desired Then
                    font.Underline = desired
                    changesMade = True
                End If
            End If

            If ff("allCaps") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("allCaps")), -1, 0)
                If font.AllCaps <> desired Then
                    font.AllCaps = desired
                    changesMade = True
                End If
            End If

            If ff("smallCaps") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("smallCaps")), -1, 0)
                If font.SmallCaps <> desired Then
                    font.SmallCaps = desired
                    changesMade = True
                End If
            End If

            If ff("strikeThrough") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("strikeThrough")), -1, 0)
                If font.StrikeThrough <> desired Then
                    font.StrikeThrough = desired
                    changesMade = True
                End If
            End If

            If ff("doubleStrikeThrough") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("doubleStrikeThrough")), -1, 0)
                If font.DoubleStrikeThrough <> desired Then
                    font.DoubleStrikeThrough = desired
                    changesMade = True
                End If
            End If

            If ff("subscript") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("subscript")), -1, 0)
                If font.Subscript <> desired Then
                    font.Subscript = desired
                    changesMade = True
                End If
            End If

            If ff("superscript") IsNot Nothing Then
                Dim desired As Integer = If(CBool(ff("superscript")), -1, 0)
                If font.Superscript <> desired Then
                    font.Superscript = desired
                    changesMade = True
                End If
            End If

            If ff("colorRGB") IsNot Nothing Then
                Dim desired As WdColor = ParseColorFromRGB(CStr(ff("colorRGB")))
                If font.Color <> desired Then
                    font.Color = desired
                    changesMade = True
                End If
            End If

            If ff("scaling") IsNot Nothing Then
                Try
                    Dim desired As Integer = CInt(ff("scaling"))
                    If font.Scaling <> desired Then
                        font.Scaling = desired
                        changesMade = True
                    End If
                Catch
                End Try
            End If

            If ff("spacing") IsNot Nothing Then
                Try
                    Dim desired As Single = CSng(ff("spacing"))
                    If Math.Abs(font.Spacing - desired) > 0.1 Then
                        font.Spacing = desired
                        changesMade = True
                    End If
                Catch
                End Try
            End If

            If ff("position") IsNot Nothing Then
                Try
                    Dim desired As Single = CSng(ff("position"))
                    If Math.Abs(font.Position - desired) > 0.1 Then
                        font.Position = desired
                        changesMade = True
                    End If
                Catch
                End Try
            End If

            If ff("kerning") IsNot Nothing Then
                Try
                    Dim desired As Single = CSng(ff("kerning"))
                    If Math.Abs(font.Kerning - desired) > 0.1 Then
                        font.Kerning = desired
                        changesMade = True
                    End If
                Catch
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"Error comparing/applying font format: {ex.Message}")
        End Try

        Return changesMade
    End Function

    ''' <summary>
    ''' Applies font format to style without comparison.
    ''' </summary>
    Private Sub ApplyFontFormatToStyle(style As Word.Style, ff As JToken)
        Dim font As Word.Font = style.Font

        Try
            If ff("name") IsNot Nothing Then font.Name = CStr(ff("name"))
            If ff("size") IsNot Nothing Then font.Size = CSng(ff("size"))
            If ff("bold") IsNot Nothing Then font.Bold = If(CBool(ff("bold")), -1, 0)
            If ff("italic") IsNot Nothing Then font.Italic = If(CBool(ff("italic")), -1, 0)
            If ff("underline") IsNot Nothing Then font.Underline = ParseUnderline(CStr(ff("underline")))
            If ff("allCaps") IsNot Nothing Then font.AllCaps = If(CBool(ff("allCaps")), -1, 0)
            If ff("smallCaps") IsNot Nothing Then font.SmallCaps = If(CBool(ff("smallCaps")), -1, 0)
            If ff("strikeThrough") IsNot Nothing Then font.StrikeThrough = If(CBool(ff("strikeThrough")), -1, 0)
            If ff("doubleStrikeThrough") IsNot Nothing Then font.DoubleStrikeThrough = If(CBool(ff("doubleStrikeThrough")), -1, 0)
            If ff("subscript") IsNot Nothing Then font.Subscript = If(CBool(ff("subscript")), -1, 0)
            If ff("superscript") IsNot Nothing Then font.Superscript = If(CBool(ff("superscript")), -1, 0)
            If ff("colorRGB") IsNot Nothing Then font.Color = ParseColorFromRGB(CStr(ff("colorRGB")))
            Try
                If ff("scaling") IsNot Nothing Then font.Scaling = CInt(ff("scaling"))
                If ff("spacing") IsNot Nothing Then font.Spacing = CSng(ff("spacing"))
                If ff("position") IsNot Nothing Then font.Position = CSng(ff("position"))
                If ff("kerning") IsNot Nothing Then font.Kerning = CSng(ff("kerning"))
            Catch
            End Try
        Catch ex As Exception
            Debug.WriteLine($"Error applying font format: {ex.Message}")
        End Try
    End Sub



    ''' <summary>
    ''' Compares and applies list format. For built-in heading styles, we ARE more aggressive
    ''' because the user explicitly wants to apply a template's numbering scheme.
    ''' Returns True if changes were made.
    ''' </summary>
    Private Function ApplyListFormatIfDifferent(doc As Word.Document, style As Word.Style, lfDef As JToken, isBuiltIn As Boolean) As Boolean
        If lfDef Is Nothing Then Return False

        Dim templateWantsList As Boolean = False
        If lfDef("hasListTemplate") IsNot Nothing Then
            templateWantsList = CBool(lfDef("hasListTemplate"))
        End If

        ' Get current list template info
        Dim currentTpl As Word.ListTemplate = Nothing
        Dim styleHasListTemplate As Boolean = False
        Try
            currentTpl = style.ListTemplate
            styleHasListTemplate = (currentTpl IsNot Nothing)
        Catch
            styleHasListTemplate = False
        End Try

        ' Case 1: Template says no list
        If Not templateWantsList Then
            If styleHasListTemplate Then
                ' Remove list from style if template says it shouldn't have one
                ' For headings, this is intentional - user wants to remove numbering
                Try
                    style.LinkToListTemplate(Nothing)
                    Debug.WriteLine($"[DocStyle] Removed list template from style '{style.NameLocal}'")
                    Return True
                Catch ex As Exception
                    Debug.WriteLine($"[DocStyle] Could not remove list template from '{style.NameLocal}': {ex.Message}")
                    Return False
                End Try
            End If
            Return False
        End If

        ' Case 2: Template wants a list - check if we need to update
        If styleHasListTemplate Then
            ' Compare fingerprints to see if the list template is actually different
            Dim currentFingerprint As String = ComputeListTemplateFingerprint(currentTpl)
            Dim desiredFingerprint As String = If(lfDef("templateFingerprint") IsNot Nothing,
                                               CStr(lfDef("templateFingerprint")), "")

            If Not String.IsNullOrEmpty(desiredFingerprint) AndAlso
           currentFingerprint = desiredFingerprint Then
                ' Same template - no changes needed
                Debug.WriteLine($"[DocStyle] List template unchanged for style '{style.NameLocal}' (fingerprints match)")
                Return False
            End If

            ' Different template - we need to replace it
            ' For heading styles especially, this is critical for numbered headings
            Debug.WriteLine($"[DocStyle] Replacing list template for style '{style.NameLocal}' (fingerprint mismatch)")
            Debug.WriteLine($"[DocStyle]   Current: {currentFingerprint}")
            Debug.WriteLine($"[DocStyle]   Desired: {desiredFingerprint}")
        End If

        ' Apply the new list template
        ApplyListFormatToStyle(doc, style, lfDef)
        Return True
    End Function

    ''' <summary>
    ''' Applies Word style definitions from template to the document.
    ''' For heading styles and styles with list templates, we are MORE aggressive
    ''' to ensure numbered headings are properly applied from the JSON template.
    ''' Shows a confirmation dialog if all styles already exist to prevent corruption.
    ''' </summary>
    Private Sub ApplyWdStyleDefinitionsToDocument(doc As Word.Document, wdStyleDefinitions As JObject)
        ' Clear shared template cache at start
        _sharedListTemplates.Clear()

        ' First pass: Check which styles already exist
        Dim allStylesExist As Boolean = True
        Dim existingStyles As New List(Of String)()
        Dim missingStyles As New List(Of String)()

        For Each prop As JProperty In wdStyleDefinitions.Properties()
            Dim styleName As String = prop.Name
            Try
                Dim style As Word.Style = doc.Styles(styleName)
                existingStyles.Add(styleName)
            Catch
                allStylesExist = False
                missingStyles.Add(styleName)
            End Try
        Next

        ' If all styles exist, ask user for confirmation before updating
        If allStylesExist AndAlso existingStyles.Count > 0 Then
            Dim confirmResult As Integer = ShowCustomYesNoBox(
            $"All {existingStyles.Count} Word styles from the template already exist in this document." & vbCrLf & vbCrLf &
            "Updating existing styles may cause formatting issues if applied multiple times (also, for the same reason, do not apply DocStyle templates twice)." & vbCrLf & vbCrLf &
            "Do you want to update the existing styles anyway?",
            "Yes, update styles", "No, skip style update", $"{AN} - Style Update")

            If confirmResult <> 1 Then
                Debug.WriteLine($"[DocStyle] User skipped style update - all {existingStyles.Count} styles already exist")
                Return
            End If

            Debug.WriteLine($"[DocStyle] User confirmed style update for {existingStyles.Count} existing styles")
        End If

        Dim stylesCreatedOrUpdated As New List(Of String)()
        Dim stylesSkipped As New List(Of String)()

        ' Sort styles by linked level so level 1 styles (which create the template) come first
        Dim sortedProps = wdStyleDefinitions.Properties().OrderBy(Function(p)
                                                                      Try
                                                                          Dim def = CType(p.Value, JObject)
                                                                          If def("listFormat") IsNot Nothing AndAlso
                                                                          def("listFormat")("linkedLevel") IsNot Nothing Then
                                                                              Return CInt(def("listFormat")("linkedLevel"))
                                                                          End If
                                                                      Catch
                                                                      End Try
                                                                      Return 0
                                                                  End Function).ToList()

        For Each prop As JProperty In sortedProps
            Dim styleName As String = prop.Name
            Dim styleDef As JObject = CType(prop.Value, JObject)

            Try
                Dim style As Word.Style = Nothing
                Dim styleExists As Boolean = False
                Dim isBuiltIn As Boolean = False

                ' Check if style exists
                Try
                    style = doc.Styles(styleName)
                    styleExists = True
                    isBuiltIn = style.BuiltIn
                Catch
                    styleExists = False
                End Try

                ' Detect if this is a heading style (built-in or custom heading)
                Dim isHeadingStyle As Boolean = False
                If Not String.IsNullOrEmpty(styleName) Then
                    isHeadingStyle = styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) OrElse
                                 styleName.StartsWith("Überschrift", StringComparison.OrdinalIgnoreCase) OrElse
                                 styleName.StartsWith("Titre", StringComparison.OrdinalIgnoreCase) OrElse
                                 styleName.StartsWith("Título", StringComparison.OrdinalIgnoreCase) OrElse
                                 styleName.StartsWith("Titolo", StringComparison.OrdinalIgnoreCase) OrElse
                                 Regex.IsMatch(styleName, "^(Heading|Überschrift|Titre|Título|Titolo)\s*\d", RegexOptions.IgnoreCase)
                End If

                ' Create style if it doesn't exist
                If Not styleExists Then
                    Try
                        style = doc.Styles.Add(styleName, WdStyleType.wdStyleTypeParagraph)
                        Debug.WriteLine($"[DocStyle] Created new wdStyle: {styleName}")

                        ' For NEW styles, apply all formatting including list
                        ApplyAllFormattingToNewStyle(doc, style, styleDef)
                        stylesCreatedOrUpdated.Add($"{styleName} (created)")

                        ' Make style visible in Quick Styles gallery
                        Try
                            style.QuickStyle = True
                        Catch
                        End Try

                        Continue For ' Skip the rest - we're done with this new style
                    Catch ex As Exception
                        Debug.WriteLine($"[DocStyle] Could not create wdStyle '{styleName}': {ex.Message}")
                        Continue For
                    End Try
                End If

                If style Is Nothing Then Continue For

                ' For EXISTING styles: update formatting
                Dim changesMade As Boolean = False

                ' Check and apply paragraph formatting differences
                If styleDef("paragraphFormat") IsNot Nothing Then
                    changesMade = ApplyParagraphFormatIfDifferent(style, styleDef("paragraphFormat")) OrElse changesMade
                End If

                ' Check and apply tab stops differences
                If styleDef("tabStops") IsNot Nothing Then
                    changesMade = ApplyTabStopsIfDifferent(style, CType(styleDef("tabStops"), JArray)) OrElse changesMade
                End If

                ' Check and apply font formatting differences
                If styleDef("fontFormat") IsNot Nothing Then
                    changesMade = ApplyFontFormatIfDifferent(style, styleDef("fontFormat")) OrElse changesMade
                End If

                ' Handle list formatting - BE AGGRESSIVE for heading styles
                If styleDef("listFormat") IsNot Nothing Then
                    Dim lfDef As JToken = styleDef("listFormat")
                    Dim templateWantsList As Boolean = False
                    If lfDef("hasListTemplate") IsNot Nothing Then
                        templateWantsList = CBool(lfDef("hasListTemplate"))
                    End If

                    Dim currentTpl As Word.ListTemplate = Nothing
                    Dim styleHasListTemplate As Boolean = False
                    Try
                        currentTpl = style.ListTemplate
                        styleHasListTemplate = (currentTpl IsNot Nothing)
                    Catch
                        styleHasListTemplate = False
                    End Try

                    If isHeadingStyle Then
                        ' HEADING STYLES: Always apply list format from template
                        ' This is critical for numbered headings (e.g., "1. Heading", "1.1 Subheading")
                        Debug.WriteLine($"[DocStyle] Processing heading style '{styleName}' (hasTemplate={styleHasListTemplate}, templateWantsList={templateWantsList})")

                        If templateWantsList Then
                            ' Template wants a list - check if we need to replace
                            Dim needsUpdate As Boolean = True

                            If styleHasListTemplate Then
                                ' Compare fingerprints
                                Dim currentFingerprint As String = ComputeListTemplateFingerprint(currentTpl)
                                Dim desiredFingerprint As String = If(lfDef("templateFingerprint") IsNot Nothing,
                                                                   CStr(lfDef("templateFingerprint")), "")

                                If Not String.IsNullOrEmpty(desiredFingerprint) AndAlso
                               Not String.IsNullOrEmpty(currentFingerprint) AndAlso
                               currentFingerprint = desiredFingerprint Then
                                    needsUpdate = False
                                    Debug.WriteLine($"[DocStyle] Heading '{styleName}' list template matches (fingerprints equal)")
                                Else
                                    Debug.WriteLine($"[DocStyle] Heading '{styleName}' list template differs:")
                                    Debug.WriteLine($"[DocStyle]   Current:  {currentFingerprint}")
                                    Debug.WriteLine($"[DocStyle]   Desired:  {desiredFingerprint}")
                                End If
                            End If

                            If needsUpdate Then
                                ' Apply new list template
                                ApplyListFormatToStyle(doc, style, lfDef)
                                changesMade = True
                                Debug.WriteLine($"[DocStyle] Applied list template to heading style '{styleName}'")
                            End If

                        ElseIf styleHasListTemplate Then
                            ' Template says NO list, but style has one - remove it
                            Try
                                style.LinkToListTemplate(Nothing)
                                changesMade = True
                                Debug.WriteLine($"[DocStyle] Removed list template from heading style '{styleName}'")
                            Catch ex As Exception
                                Debug.WriteLine($"[DocStyle] Could not remove list from '{styleName}': {ex.Message}")
                            End Try
                        End If

                    Else
                        ' NON-HEADING STYLES: Be more conservative but still allow updates
                        If templateWantsList Then
                            If styleHasListTemplate Then
                                ' Compare fingerprints - only update if different
                                Dim currentFingerprint As String = ComputeListTemplateFingerprint(currentTpl)
                                Dim desiredFingerprint As String = If(lfDef("templateFingerprint") IsNot Nothing,
                                                                   CStr(lfDef("templateFingerprint")), "")

                                If String.IsNullOrEmpty(desiredFingerprint) OrElse
                               String.IsNullOrEmpty(currentFingerprint) OrElse
                               currentFingerprint <> desiredFingerprint Then
                                    ' Different or unknown - update it
                                    ApplyListFormatToStyle(doc, style, lfDef)
                                    changesMade = True
                                    Debug.WriteLine($"[DocStyle] Updated list template for style '{styleName}' (fingerprint mismatch)")
                                Else
                                    Debug.WriteLine($"[DocStyle] List template unchanged for style '{styleName}' (fingerprints match)")
                                End If
                            Else
                                ' No existing list - add one (even for built-in if template specifies it)
                                ApplyListFormatToStyle(doc, style, lfDef)
                                changesMade = True
                                Debug.WriteLine($"[DocStyle] Added list template to style '{styleName}'")
                            End If

                        ElseIf styleHasListTemplate AndAlso Not isBuiltIn Then
                            ' Template says no list, style has one, and it's not built-in - remove
                            Try
                                style.LinkToListTemplate(Nothing)
                                changesMade = True
                                Debug.WriteLine($"[DocStyle] Removed list template from style '{styleName}'")
                            Catch ex As Exception
                                Debug.WriteLine($"[DocStyle] Could not remove list from '{styleName}': {ex.Message}")
                            End Try
                        End If
                    End If
                End If

                If changesMade Then
                    stylesCreatedOrUpdated.Add($"{styleName} (updated)")
                    Debug.WriteLine($"[DocStyle] Updated existing wdStyle: {styleName} (BuiltIn: {isBuiltIn}, Heading: {isHeadingStyle})")
                Else
                    stylesSkipped.Add(styleName)
                    Debug.WriteLine($"[DocStyle] Skipped wdStyle (no changes needed): {styleName}")
                End If

                ' Make style visible in Quick Styles gallery
                Try
                    style.QuickStyle = True
                Catch
                End Try

            Catch ex As Exception
                Debug.WriteLine($"[DocStyle] Error creating/updating wdStyle '{styleName}': {ex.Message}")
            End Try
        Next

        ' Clear cache after processing
        _sharedListTemplates.Clear()

        If stylesCreatedOrUpdated.Count > 0 Then
            Debug.WriteLine($"[DocStyle] Word styles processed: {String.Join(", ", stylesCreatedOrUpdated)}")
        End If
        If stylesSkipped.Count > 0 Then
            Debug.WriteLine($"[DocStyle] Word styles unchanged: {String.Join(", ", stylesSkipped)}")
        End If
    End Sub


    ''' <summary>
    ''' Applies user style formatting (paragraph, font, list) from the template to a paragraph.
    ''' This applies the formatting captured in the template as overrides after the Word style.
    ''' </summary>
    Private Sub ApplyUserStyleFormattingFromTemplate(para As Word.Paragraph, userStyleDef As JObject)
        If userStyleDef Is Nothing Then Return

        Try
            ' Keep a reference so we can reapply after list operations (RemoveNumbers can clobber indents/tabs).
            Dim pf As JToken = userStyleDef("paragraphFormatting")

            ' 1) Apply paragraph formatting overrides (initial pass)
            If pf IsNot Nothing Then
                If pf("alignment") IsNot Nothing Then para.Alignment = ParseAlignment(CStr(pf("alignment")))
                If pf("leftIndent") IsNot Nothing Then para.LeftIndent = CSng(pf("leftIndent"))
                If pf("rightIndent") IsNot Nothing Then para.RightIndent = CSng(pf("rightIndent"))
                If pf("firstLineIndent") IsNot Nothing Then para.FirstLineIndent = CSng(pf("firstLineIndent"))
                If pf("spaceBefore") IsNot Nothing Then para.SpaceBefore = CSng(pf("spaceBefore"))
                If pf("spaceAfter") IsNot Nothing Then para.SpaceAfter = CSng(pf("spaceAfter"))
                If pf("lineSpacing") IsNot Nothing Then para.LineSpacing = CSng(pf("lineSpacing"))
                If pf("lineSpacingRule") IsNot Nothing Then para.Format.LineSpacingRule = ParseLineSpacingRule(CStr(pf("lineSpacingRule")))
                If pf("keepTogether") IsNot Nothing Then para.KeepTogether = CInt(pf("keepTogether"))
                If pf("keepWithNext") IsNot Nothing Then para.KeepWithNext = CInt(pf("keepWithNext"))
            End If

            ' 2) Apply font formatting overrides
            If userStyleDef("fontFormatting") IsNot Nothing Then
                Dim ff = userStyleDef("fontFormatting")
                Dim rng As Word.Range = para.Range
                If ff("fontName") IsNot Nothing AndAlso CStr(ff("fontName")) <> "mixed" Then rng.Font.Name = CStr(ff("fontName"))
                If ff("fontSize") IsNot Nothing AndAlso CStr(ff("fontSize")) <> "mixed" Then rng.Font.Size = CSng(ff("fontSize"))
                If ff("bold") IsNot Nothing AndAlso TypeOf ff("bold") Is JValue AndAlso ff("bold").Type = JTokenType.Boolean Then rng.Font.Bold = If(CBool(ff("bold")), -1, 0)
                If ff("italic") IsNot Nothing AndAlso TypeOf ff("italic") Is JValue AndAlso ff("italic").Type = JTokenType.Boolean Then rng.Font.Italic = If(CBool(ff("italic")), -1, 0)
                If ff("allCaps") IsNot Nothing AndAlso TypeOf ff("allCaps") Is JValue AndAlso ff("allCaps").Type = JTokenType.Boolean Then rng.Font.AllCaps = If(CBool(ff("allCaps")), -1, 0)
                If ff("smallCaps") IsNot Nothing AndAlso TypeOf ff("smallCaps") Is JValue AndAlso ff("smallCaps").Type = JTokenType.Boolean Then rng.Font.SmallCaps = If(CBool(ff("smallCaps")), -1, 0)
                If ff("underline") IsNot Nothing Then rng.Font.Underline = ParseUnderline(CStr(ff("underline")))
            End If

            ' 3) Handle list formatting (RemoveNumbers may destroy indents/tabs)
            Dim didRemoveNumbers As Boolean = False
            If userStyleDef("listFormatting") IsNot Nothing Then
                Dim lf = userStyleDef("listFormatting")
                Dim templateHasList As Boolean = If(lf("hasList") IsNot Nothing, CBool(lf("hasList")), False)

                Dim currentHasList As Boolean = False
                Try
                    currentHasList = (para.Range.ListFormat.ListType <> WdListType.wdListNoNumbering)
                Catch
                End Try

                If Not templateHasList AndAlso currentHasList Then
                    para.Range.ListFormat.RemoveNumbers()
                    didRemoveNumbers = True
                ElseIf templateHasList AndAlso Not currentHasList Then
                    Dim listType As String = If(lf("listType") IsNot Nothing, CStr(lf("listType")), "")
                    If listType.Contains("Bullet") Then
                        para.Range.ListFormat.ApplyBulletDefault()
                    ElseIf listType.Contains("Outline") Then
                        para.Range.ListFormat.ApplyOutlineNumberDefault()
                    ElseIf listType.Contains("Number") Then
                        para.Range.ListFormat.ApplyNumberDefault()
                    End If
                End If
            End If

            ' 4) Re-apply paragraph indents AFTER RemoveNumbers (this is the bug fix for MAIN TITLE)
            If didRemoveNumbers AndAlso pf IsNot Nothing Then
                If pf("leftIndent") IsNot Nothing Then para.LeftIndent = CSng(pf("leftIndent"))
                If pf("rightIndent") IsNot Nothing Then para.RightIndent = CSng(pf("rightIndent"))
                If pf("firstLineIndent") IsNot Nothing Then para.FirstLineIndent = CSng(pf("firstLineIndent"))
            End If

            ' 5) Apply tab stops if specified (RemoveNumbers can also affect tabs)
            If userStyleDef("tabStops") IsNot Nothing Then
                Try
                    para.TabStops.ClearAll()
                    For Each tabDef As JObject In CType(userStyleDef("tabStops"), JArray)
                        Dim position As Single = CSng(tabDef("pos"))
                        Dim alignStr As String = If(tabDef("align") IsNot Nothing, CStr(tabDef("align")), "Left")
                        Dim alignment As WdTabAlignment = ParseTabAlignment("wdAlignTab" & alignStr)
                        Dim leader As WdTabLeader = WdTabLeader.wdTabLeaderSpaces
                        If tabDef("leader") IsNot Nothing Then
                            leader = ParseTabLeader("wdTabLeader" & CStr(tabDef("leader")))
                        End If
                        para.TabStops.Add(position, alignment, leader)
                    Next
                Catch
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"Error applying user style formatting: {ex.Message}")
        End Try
    End Sub


    ''' <summary>
    ''' Applies formatting specification to a paragraph with list preservation support.
    ''' </summary>
    Private Sub ApplyFormattingToParagraphEx(doc As Word.Document, para As Word.Paragraph,
                                              wdStyleName As String, formatting As JToken,
                                              preserveList As Boolean,
                                              originalListType As WdListType,
                                              originalListLevel As Integer)
        ' First apply style if specified
        If Not String.IsNullOrWhiteSpace(wdStyleName) Then
            Try
                para.Style = doc.Styles(wdStyleName)
            Catch
                ' Style doesn't exist, try Normal
                Try
                    para.Style = doc.Styles("Normal")
                Catch
                End Try
            End Try
        End If

        ' Restore list formatting if needed (applying a style may remove it)
        If preserveList AndAlso originalListType <> WdListType.wdListNoNumbering Then
            Try
                ' Check if list was removed by style application
                If para.Range.ListFormat.ListType = WdListType.wdListNoNumbering Then
                    If originalListType = WdListType.wdListBullet Then
                        para.Range.ListFormat.ApplyBulletDefault()
                    ElseIf originalListType = WdListType.wdListSimpleNumbering OrElse
                           originalListType = WdListType.wdListListNumOnly OrElse
                           originalListType = WdListType.wdListMixedNumbering Then
                        para.Range.ListFormat.ApplyNumberDefault()
                    ElseIf originalListType = WdListType.wdListOutlineNumbering Then
                        para.Range.ListFormat.ApplyOutlineNumberDefault()
                    End If

                    ' Restore level
                    If originalListLevel > 1 Then
                        For lvl As Integer = 2 To originalListLevel
                            para.Range.ListFormat.ListIndent()
                        Next
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"Could not restore list formatting: {ex.Message}")
            End Try
        End If

        ' Then apply additional formatting overrides from LLM response
        If formatting Is Nothing Then Return

        Try
            ' Paragraph formatting
            If formatting("alignment") IsNot Nothing AndAlso CStr(formatting("alignment")) <> "" Then
                para.Alignment = ParseAlignment(CStr(formatting("alignment")))
            End If
            If formatting("leftIndent") IsNot Nothing Then para.LeftIndent = CSng(formatting("leftIndent"))
            If formatting("rightIndent") IsNot Nothing Then para.RightIndent = CSng(formatting("rightIndent"))
            If formatting("firstLineIndent") IsNot Nothing Then para.FirstLineIndent = CSng(formatting("firstLineIndent"))
            If formatting("spaceBefore") IsNot Nothing Then para.SpaceBefore = CSng(formatting("spaceBefore"))
            If formatting("spaceAfter") IsNot Nothing Then para.SpaceAfter = CSng(formatting("spaceAfter"))
            If formatting("lineSpacing") IsNot Nothing Then para.LineSpacing = CSng(formatting("lineSpacing"))

            ' Font formatting
            Dim rng As Word.Range = para.Range
            If formatting("fontName") IsNot Nothing AndAlso CStr(formatting("fontName")) <> "" Then
                rng.Font.Name = CStr(formatting("fontName"))
            End If
            If formatting("fontSize") IsNot Nothing Then rng.Font.Size = CSng(formatting("fontSize"))
            If formatting("bold") IsNot Nothing Then rng.Font.Bold = If(CBool(formatting("bold")), -1, 0)
            If formatting("italic") IsNot Nothing Then rng.Font.Italic = If(CBool(formatting("italic")), -1, 0)
            If formatting("underline") IsNot Nothing Then rng.Font.Underline = ParseUnderline(CStr(formatting("underline")))
        Catch
        End Try
    End Sub



    ''' <summary>
    ''' Applies list formatting to a style, correctly handling:
    ''' - Shared outline templates (multiple styles using same template at different levels)
    ''' - Bullet characters with correct fonts
    ''' - All level properties including tab positions
    ''' - Separate caching for outline vs non-outline templates
    '''
    ''' PATCH: Avoid dependency on ListGalleries for bullets (non-deterministic), use doc.ListTemplates.Add.
    ''' </summary>
    Private Sub ApplyListFormatToStyle(doc As Word.Document, style As Word.Style, listFormatDef As JToken)
        If listFormatDef Is Nothing Then Return
        If listFormatDef("hasListTemplate") Is Nothing OrElse Not CBool(listFormatDef("hasListTemplate")) Then Return

        Try
            Dim isOutline As Boolean = If(listFormatDef("outlineNumbered") IsNot Nothing, CBool(listFormatDef("outlineNumbered")), False)
            Dim linkedLevel As Integer = If(listFormatDef("linkedLevel") IsNot Nothing, CInt(listFormatDef("linkedLevel")), 1)
            Dim fingerprint As String = If(listFormatDef("templateFingerprint") IsNot Nothing, CStr(listFormatDef("templateFingerprint")), "")

            ' Check if this is a bullet-only template (all levels are bullets)
            Dim isBulletTemplate As Boolean = False
            If listFormatDef("levels") IsNot Nothing Then
                isBulletTemplate = True
                For Each levelDef As JObject In CType(listFormatDef("levels"), JArray)
                    Dim ns As String = If(levelDef("numberStyle") IsNot Nothing, CStr(levelDef("numberStyle")), "")
                    If Not ns.ToLower().Contains("bullet") Then
                        isBulletTemplate = False
                        Exit For
                    End If
                Next
            End If

            ' Include outline flag AND bullet flag in cache key
            Dim cacheKey As String = $"{If(isOutline, "O", "N")}:{If(isBulletTemplate, "B", "N")}:{fingerprint}"

            Dim listTemplate As Word.ListTemplate = Nothing

            ' Check if we already created this template
            If Not String.IsNullOrEmpty(fingerprint) Then
                If _sharedListTemplates.ContainsKey(cacheKey) Then
                    listTemplate = _sharedListTemplates(cacheKey)
                    Debug.WriteLine($"Reusing shared list template for style '{style.NameLocal}' at level {linkedLevel}")
                End If
            End If

            ' Create new template if needed
            If listTemplate Is Nothing Then
                ' PATCH: Always create our own template for bullets too (gallery templates vary by user/profile)
                listTemplate = doc.ListTemplates.Add(OutlineNumbered:=isOutline)

                Debug.WriteLine($"Created new list template for style '{style.NameLocal}' (outline={isOutline}, bullet={isBulletTemplate}, levels={listTemplate.ListLevels.Count})")

                If listFormatDef("levels") IsNot Nothing Then
                    For Each levelDef As JObject In CType(listFormatDef("levels"), JArray)
                        Try
                            Dim levelNum As Integer = CInt(levelDef("level"))
                            If levelNum > listTemplate.ListLevels.Count Then Continue For

                            Dim level As Word.ListLevel = listTemplate.ListLevels(levelNum)

                            Dim numberStyleStr As String = If(levelDef("numberStyle") IsNot Nothing, CStr(levelDef("numberStyle")), "")
                            Dim isBulletLevel As Boolean = numberStyleStr.ToLower().Contains("bullet")

                            If isBulletLevel OrElse isBulletTemplate Then
                                ' Bullet config
                                Dim bulletFontName As String = If(levelDef("bulletFont") IsNot Nothing, CStr(levelDef("bulletFont")), "Symbol")
                                Dim bulletCharCode As Integer = If(levelDef("bulletCharCode") IsNot Nothing, CInt(levelDef("bulletCharCode")), &H2022)

                                ' IMPORTANT: reset first, then assign deterministically
                                Try : level.Font.Reset() : Catch : End Try
                                Try : level.Font.Name = bulletFontName : Catch : End Try
                                Try : level.NumberStyle = WdListNumberStyle.wdListNumberStyleBullet : Catch : End Try
                                Try : level.NumberFormat = ChrW(bulletCharCode) : Catch : End Try
                            Else
                                ' Numbered config
                                Dim numberStyle As WdListNumberStyle = ParseNumberStyle(If(levelDef("numberStyle") IsNot Nothing, CStr(levelDef("numberStyle")), "wdListNumberStyleArabic"))
                                Try : level.NumberStyle = numberStyle : Catch : End Try
                                If levelDef("numberFormat") IsNot Nothing AndAlso Not String.IsNullOrEmpty(CStr(levelDef("numberFormat"))) Then
                                    Try : level.NumberFormat = CStr(levelDef("numberFormat")) : Catch : End Try
                                End If
                            End If

                            ' Common positions
                            If levelDef("numberPosition") IsNot Nothing Then level.NumberPosition = CSng(levelDef("numberPosition"))
                            If levelDef("textPosition") IsNot Nothing Then level.TextPosition = CSng(levelDef("textPosition"))
                            If levelDef("tabPosition") IsNot Nothing Then level.TabPosition = CSng(levelDef("tabPosition"))
                            If levelDef("startAt") IsNot Nothing Then level.StartAt = CInt(levelDef("startAt"))
                            If levelDef("alignment") IsNot Nothing Then level.Alignment = ParseListLevelAlignment(CStr(levelDef("alignment")))

                            If levelDef("trailingCharacter") IsNot Nothing Then
                                Try
                                    level.TrailingCharacter = ParseTrailingCharacter(CStr(levelDef("trailingCharacter")))
                                Catch
                                End Try
                            End If

                        Catch ex As Exception
                            Debug.WriteLine($"Error configuring list level: {ex.Message}")
                        End Try
                    Next
                End If

                ' Cache templates for reuse
                If Not String.IsNullOrEmpty(fingerprint) AndAlso listTemplate IsNot Nothing Then
                    _sharedListTemplates(cacheKey) = listTemplate
                    Debug.WriteLine($"Cached list template for style '{style.NameLocal}' (key={cacheKey})")
                End If
            End If

            ' Link template to style at the CORRECT level
            If listTemplate IsNot Nothing Then
                style.LinkToListTemplate(listTemplate, linkedLevel)
                Debug.WriteLine($"Linked list template to style '{style.NameLocal}' at level {linkedLevel}")
            End If

        Catch ex As Exception
            Debug.WriteLine($"Error applying list template to style '{style.NameLocal}': {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Parses line spacing rule string to Word enum.
    ''' </summary>
    Private Function ParseLineSpacingRule(rule As String) As WdLineSpacing
        If String.IsNullOrWhiteSpace(rule) Then Return WdLineSpacing.wdLineSpaceSingle

        Select Case rule.ToLower().Replace("wdlinespace", "").Replace("_", "")
            Case "single" : Return WdLineSpacing.wdLineSpaceSingle
            Case "1pt5" : Return WdLineSpacing.wdLineSpace1pt5
            Case "double" : Return WdLineSpacing.wdLineSpaceDouble
            Case "atleast" : Return WdLineSpacing.wdLineSpaceAtLeast
            Case "exactly" : Return WdLineSpacing.wdLineSpaceExactly
            Case "multiple" : Return WdLineSpacing.wdLineSpaceMultiple
            Case Else : Return WdLineSpacing.wdLineSpaceSingle
        End Select
    End Function

    ''' <summary>
    ''' Parses outline level string to Word enum.
    ''' </summary>
    Private Function ParseOutlineLevel(level As String) As WdOutlineLevel
        If String.IsNullOrWhiteSpace(level) Then Return WdOutlineLevel.wdOutlineLevelBodyText

        Select Case level.ToLower().Replace("wdoutlinelevel", "").Replace("_", "")
            Case "1" : Return WdOutlineLevel.wdOutlineLevel1
            Case "2" : Return WdOutlineLevel.wdOutlineLevel2
            Case "3" : Return WdOutlineLevel.wdOutlineLevel3
            Case "4" : Return WdOutlineLevel.wdOutlineLevel4
            Case "5" : Return WdOutlineLevel.wdOutlineLevel5
            Case "6" : Return WdOutlineLevel.wdOutlineLevel6
            Case "7" : Return WdOutlineLevel.wdOutlineLevel7
            Case "8" : Return WdOutlineLevel.wdOutlineLevel8
            Case "9" : Return WdOutlineLevel.wdOutlineLevel9
            Case "bodytext" : Return WdOutlineLevel.wdOutlineLevelBodyText
            Case Else : Return WdOutlineLevel.wdOutlineLevelBodyText
        End Select
    End Function

    ''' <summary>
    ''' Parses tab alignment string to Word enum.
    ''' </summary>
    Private Function ParseTabAlignment(alignment As String) As WdTabAlignment
        If String.IsNullOrWhiteSpace(alignment) Then Return WdTabAlignment.wdAlignTabLeft

        Select Case alignment.ToLower().Replace("wdaligntab", "").Replace("_", "")
            Case "left" : Return WdTabAlignment.wdAlignTabLeft
            Case "center" : Return WdTabAlignment.wdAlignTabCenter
            Case "right" : Return WdTabAlignment.wdAlignTabRight
            Case "decimal" : Return WdTabAlignment.wdAlignTabDecimal
            Case "bar" : Return WdTabAlignment.wdAlignTabBar
            Case "list" : Return WdTabAlignment.wdAlignTabList
            Case Else : Return WdTabAlignment.wdAlignTabLeft
        End Select
    End Function

    ''' <summary>
    ''' Parses tab leader string to Word enum.
    ''' </summary>
    Private Function ParseTabLeader(leader As String) As WdTabLeader
        If String.IsNullOrWhiteSpace(leader) Then Return WdTabLeader.wdTabLeaderSpaces

        Select Case leader.ToLower().Replace("wdtableader", "").Replace("_", "")
            Case "spaces" : Return WdTabLeader.wdTabLeaderSpaces
            Case "dots" : Return WdTabLeader.wdTabLeaderDots
            Case "dashes" : Return WdTabLeader.wdTabLeaderDashes
            Case "lines" : Return WdTabLeader.wdTabLeaderLines
            Case "heavy" : Return WdTabLeader.wdTabLeaderHeavy
            Case "middledot" : Return WdTabLeader.wdTabLeaderMiddleDot
            Case Else : Return WdTabLeader.wdTabLeaderSpaces
        End Select
    End Function

    ''' <summary>
    ''' Parses underline string to Word enum.
    ''' </summary>
    Private Function ParseUnderline(underline As String) As WdUnderline
        If String.IsNullOrWhiteSpace(underline) Then Return WdUnderline.wdUnderlineNone

        Select Case underline.ToLower().Replace("wdunderline", "").Replace("_", "")
            Case "none" : Return WdUnderline.wdUnderlineNone
            Case "single" : Return WdUnderline.wdUnderlineSingle
            Case "words" : Return WdUnderline.wdUnderlineWords
            Case "double" : Return WdUnderline.wdUnderlineDouble
            Case "dotted" : Return WdUnderline.wdUnderlineDotted
            Case "thick" : Return WdUnderline.wdUnderlineThick
            Case "dash" : Return WdUnderline.wdUnderlineDash
            Case "dotdash" : Return WdUnderline.wdUnderlineDotDash
            Case "dotdotdash" : Return WdUnderline.wdUnderlineDotDotDash
            Case "wavy" : Return WdUnderline.wdUnderlineWavy
            Case "wavyheavy" : Return WdUnderline.wdUnderlineWavyHeavy
            Case "wavydouble" : Return WdUnderline.wdUnderlineWavyDouble
            Case "dashlong" : Return WdUnderline.wdUnderlineDashLong
            Case "dashheavy" : Return WdUnderline.wdUnderlineDashHeavy
            Case "dotdashheavy" : Return WdUnderline.wdUnderlineDotDashHeavy
            Case "dotdotdashheavy" : Return WdUnderline.wdUnderlineDotDotDashHeavy
            Case "dashlongheavy" : Return WdUnderline.wdUnderlineDashLongHeavy
            Case Else : Return WdUnderline.wdUnderlineNone
        End Select
    End Function

    ''' <summary>
    ''' Parses RGB hex string to Word color.
    ''' </summary>
    Private Function ParseColorFromRGB(rgb As String) As WdColor
        If String.IsNullOrWhiteSpace(rgb) OrElse rgb = "auto" OrElse rgb = "unknown" Then
            Return WdColor.wdColorAutomatic
        End If

        Try
            rgb = rgb.TrimStart("#"c)
            If rgb.Length = 6 Then
                Dim r As Integer = System.Convert.ToInt32(rgb.Substring(0, 2), 16)
                Dim g As Integer = System.Convert.ToInt32(rgb.Substring(2, 2), 16)
                Dim b As Integer = System.Convert.ToInt32(rgb.Substring(4, 2), 16)
                Return CType(r + (g * 256) + (b * 65536), WdColor)
            End If
        Catch
        End Try

        Return WdColor.wdColorAutomatic
    End Function


    ''' <summary>
    ''' Creates a simple bullet template as fallback when gallery approach fails.
    ''' </summary>
    Private Function CreateSimpleBulletTemplate(doc As Word.Document, listFormatDef As JToken) As Word.ListTemplate
        Try
            ' Create a temporary range to apply bullet format
            Dim tempRange As Word.Range = doc.Content
            tempRange.Collapse(WdCollapseDirection.wdCollapseEnd)

            ' Apply default bullet
            tempRange.ListFormat.ApplyBulletDefault()

            ' Get the template that was created
            Dim listTemplate As Word.ListTemplate = tempRange.ListFormat.ListTemplate

            ' Remove the bullet from the temp range
            tempRange.ListFormat.RemoveNumbers()

            ' Now customize the template if we have level definitions
            If listFormatDef("levels") IsNot Nothing Then
                Dim firstLevel As JObject = CType(listFormatDef("levels"), JArray).FirstOrDefault()
                If firstLevel IsNot Nothing Then
                    Try
                        Dim level As Word.ListLevel = listTemplate.ListLevels(1)

                        Dim bulletFontName As String = If(firstLevel("bulletFont") IsNot Nothing, CStr(firstLevel("bulletFont")), "Symbol")
                        Dim bulletCharCode As Integer = If(firstLevel("bulletCharCode") IsNot Nothing, CInt(firstLevel("bulletCharCode")), 61623)

                        level.Font.Name = bulletFontName

                        If firstLevel("numberPosition") IsNot Nothing Then level.NumberPosition = CSng(firstLevel("numberPosition"))
                        If firstLevel("textPosition") IsNot Nothing Then level.TextPosition = CSng(firstLevel("textPosition"))
                        If firstLevel("tabPosition") IsNot Nothing Then level.TabPosition = CSng(firstLevel("tabPosition"))
                    Catch
                    End Try
                End If
            End If

            Debug.WriteLine("Created simple bullet template as fallback")
            Return listTemplate
        Catch ex As Exception
            Debug.WriteLine($"Failed to create simple bullet template: {ex.Message}")
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' Parses list level alignment string to Word enum.
    ''' </summary>
    Private Function ParseListLevelAlignment(alignment As String) As WdListLevelAlignment
        If String.IsNullOrWhiteSpace(alignment) Then Return WdListLevelAlignment.wdListLevelAlignLeft

        Select Case alignment.ToLower().Replace("wdlistlevelalign", "").Replace("_", "")
            Case "left" : Return WdListLevelAlignment.wdListLevelAlignLeft
            Case "center" : Return WdListLevelAlignment.wdListLevelAlignCenter
            Case "right" : Return WdListLevelAlignment.wdListLevelAlignRight
            Case Else : Return WdListLevelAlignment.wdListLevelAlignLeft
        End Select
    End Function

    ''' <summary>
    ''' Parses trailing character string to Word enum.
    ''' </summary>
    Private Function ParseTrailingCharacter(trailing As String) As WdTrailingCharacter
        If String.IsNullOrWhiteSpace(trailing) Then Return WdTrailingCharacter.wdTrailingTab

        Select Case trailing.ToLower().Replace("wdtrailing", "").Replace("_", "")
            Case "tab" : Return WdTrailingCharacter.wdTrailingTab
            Case "space" : Return WdTrailingCharacter.wdTrailingSpace
            Case "none" : Return WdTrailingCharacter.wdTrailingNone
            Case Else : Return WdTrailingCharacter.wdTrailingTab
        End Select
    End Function
    ''' <summary>
    ''' Parses number style string to Word enum.
    ''' </summary>
    Private Function ParseNumberStyle(style As String) As WdListNumberStyle
        If String.IsNullOrWhiteSpace(style) Then Return WdListNumberStyle.wdListNumberStyleArabic

        Select Case style.ToLower().Replace("wdlistnumberstyle", "").Replace("_", "")
            Case "arabic" : Return WdListNumberStyle.wdListNumberStyleArabic
            Case "uppercaseroman" : Return WdListNumberStyle.wdListNumberStyleUppercaseRoman
            Case "lowercaseroman" : Return WdListNumberStyle.wdListNumberStyleLowercaseRoman
            Case "uppercaseletter" : Return WdListNumberStyle.wdListNumberStyleUppercaseLetter
            Case "lowercaseletter" : Return WdListNumberStyle.wdListNumberStyleLowercaseLetter
            Case "bullet" : Return WdListNumberStyle.wdListNumberStyleBullet
            Case "none" : Return WdListNumberStyle.wdListNumberStyleNone
            Case "ordinal" : Return WdListNumberStyle.wdListNumberStyleOrdinal
            Case "ordinaltext" : Return WdListNumberStyle.wdListNumberStyleOrdinalText
            Case "cardinaltext" : Return WdListNumberStyle.wdListNumberStyleCardinalText
            Case "legal" : Return WdListNumberStyle.wdListNumberStyleLegal
            Case "legalllz" : Return WdListNumberStyle.wdListNumberStyleLegalLZ
            Case "arabicfullwidth" : Return WdListNumberStyle.wdListNumberStyleArabicFullWidth
            Case "arabicllz" : Return WdListNumberStyle.wdListNumberStyleArabicLZ
            Case Else : Return WdListNumberStyle.wdListNumberStyleArabic
        End Select
    End Function

#End Region

#Region "Helper Classes and Functions"

    ''' <summary>
    ''' Represents a style template file.
    ''' </summary>
    Private Class DocStyleTemplate
        Public Property DisplayName As String
        Public Property FilePath As String
        Public Property IsLocal As Boolean
    End Class

    ''' <summary>
    ''' Loads available style templates from configured paths, reading display names from JSON.
    ''' </summary>
    Private Function LoadStyleTemplates(pathGlobal As String, pathLocal As String) As List(Of DocStyleTemplate)
        Dim templates As New List(Of DocStyleTemplate)()

        ' Load from global path
        If Not String.IsNullOrWhiteSpace(pathGlobal) AndAlso Directory.Exists(pathGlobal) Then
            Try
                For Each f In Directory.GetFiles(pathGlobal, $"{AN2}-ds-*.json", SearchOption.TopDirectoryOnly)
                    Dim template As DocStyleTemplate = TryLoadTemplateMetadata(f, False)
                    If template IsNot Nothing Then
                        templates.Add(template)
                    End If
                Next
            Catch
            End Try
        End If

        ' Load from local path
        If Not String.IsNullOrWhiteSpace(pathLocal) AndAlso Directory.Exists(pathLocal) Then
            Try
                For Each f In Directory.GetFiles(pathLocal, $"{AN2}-ds-*.json", SearchOption.TopDirectoryOnly)
                    Dim template As DocStyleTemplate = TryLoadTemplateMetadata(f, True)
                    If template IsNot Nothing Then
                        templates.Add(template)
                    End If
                Next
            Catch
            End Try
        End If

        Return templates
    End Function

    ''' <summary>
    ''' Attempts to load template metadata from a JSON file.
    ''' Returns Nothing if the file cannot be parsed as valid JSON.
    ''' </summary>
    Private Function TryLoadTemplateMetadata(filePath As String, isLocal As Boolean) As DocStyleTemplate
        Try
            Dim jsonContent As String = File.ReadAllText(filePath, Encoding.UTF8)
            Dim jsonObj As JObject = JObject.Parse(jsonContent)

            ' Get display name from JSON, fall back to filename
            Dim displayName As String = ""
            If jsonObj("templateName") IsNot Nothing Then
                displayName = jsonObj("templateName").ToString()
            End If

            If String.IsNullOrWhiteSpace(displayName) Then
                displayName = Path.GetFileNameWithoutExtension(filePath).Replace($"{AN2}-ds-", "")
            End If

            Return New DocStyleTemplate With {
                .DisplayName = displayName,
                .FilePath = filePath,
                .IsLocal = isLocal
            }
        Catch
            ' Invalid JSON or unreadable file - skip it
            Debug.WriteLine($"Skipping invalid template file: {filePath}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Extracts JSON from LLM response (handles markdown code blocks).
    ''' </summary>
    Private Function ExtractJsonFromResponse(response As String) As String
        If String.IsNullOrWhiteSpace(response) Then Return "{}"

        ' Try to extract from markdown code block
        Dim match As Match = Regex.Match(response, "```(?:json)?\s*([\s\S]*?)```", RegexOptions.IgnoreCase)
        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        ' Try to find JSON array or object
        Dim startIdx As Integer = response.IndexOfAny(New Char() {"["c, "{"c})
        If startIdx >= 0 Then
            Return response.Substring(startIdx).Trim()
        End If

        Return response.Trim()
    End Function

    ''' <summary>
    ''' Parses alignment string to Word alignment enum.
    ''' </summary>
    Private Function ParseAlignment(alignment As String) As WdParagraphAlignment
        If String.IsNullOrWhiteSpace(alignment) Then Return WdParagraphAlignment.wdAlignParagraphLeft

        Select Case alignment.ToLower().Replace("wdalignparagraph", "").Replace("_", "")
            Case "left" : Return WdParagraphAlignment.wdAlignParagraphLeft
            Case "center" : Return WdParagraphAlignment.wdAlignParagraphCenter
            Case "right" : Return WdParagraphAlignment.wdAlignParagraphRight
            Case "justify" : Return WdParagraphAlignment.wdAlignParagraphJustify
            Case Else : Return WdParagraphAlignment.wdAlignParagraphLeft
        End Select
    End Function

#End Region

#Region "Existing Helper Functions (from original code)"

    Private Function ExtractParagraphFormat(para As Word.Paragraph, paraRange As Word.Range) As JObject
        Dim pf As New JObject()
        Try
            pf("alignment") = para.Alignment.ToString()
            pf("leftIndent") = para.LeftIndent
            pf("rightIndent") = para.RightIndent
            pf("firstLineIndent") = para.FirstLineIndent
            pf("spaceBefore") = para.SpaceBefore
            pf("spaceAfter") = para.SpaceAfter
            pf("spaceBeforeAuto") = para.SpaceBeforeAuto
            pf("spaceAfterAuto") = para.SpaceAfterAuto
            pf("lineSpacing") = para.LineSpacing
            pf("lineSpacingRule") = para.LineSpacingRule.ToString()
            pf("keepTogether") = para.KeepTogether
            pf("keepWithNext") = para.KeepWithNext
            pf("pageBreakBefore") = para.PageBreakBefore
            pf("widowControl") = para.WidowControl
            pf("outlineLevel") = para.OutlineLevel.ToString()

            Try : pf("hyphenation") = para.Hyphenation : Catch : pf("hyphenation") = "default" : End Try

            Try
                Dim paraFormat As Object = para.Format
                pf("noSpaceBetweenParagraphsOfSameStyle") = paraFormat.NoSpaceBetweenParagraphsOfSameStyle
            Catch : End Try

            Try
                Dim paraFormat As Object = para.Format
                pf("mirrorIndents") = paraFormat.MirrorIndents
            Catch : End Try

            Try : pf("readingOrder") = para.ReadingOrder.ToString() : Catch : End Try

            Try
                Dim paraFormat As Object = para.Format
                pf("contextualSpacing") = paraFormat.ContextualSpacing
            Catch : End Try

        Catch ex As Exception
            pf("error") = ex.Message
        End Try
        Return pf
    End Function

    Private Function ExtractFontFormat(rng As Word.Range) As JObject
        Dim ff As New JObject()
        Try
            Dim font As Word.Font = rng.Font

            If font.Name <> CStr(WdConstants.wdUndefined) Then ff("fontName") = font.Name Else ff("fontName") = "mixed"
            If font.Size <> CSng(WdConstants.wdUndefined) Then ff("fontSize") = font.Size Else ff("fontSize") = "mixed"
            If font.Bold <> CInt(WdConstants.wdUndefined) Then ff("bold") = (font.Bold = -1) Else ff("bold") = "mixed"
            If font.Italic <> CInt(WdConstants.wdUndefined) Then ff("italic") = (font.Italic = -1) Else ff("italic") = "mixed"
            If font.Underline <> CType(WdConstants.wdUndefined, WdUnderline) Then ff("underline") = font.Underline.ToString() Else ff("underline") = "mixed"

            Try
                If font.UnderlineColor <> CType(WdConstants.wdUndefined, WdColor) Then
                    ff("underlineColor") = font.UnderlineColor.ToString()
                    ff("underlineColorRGB") = ColorToRGB(font.UnderlineColor)
                End If
            Catch : End Try

            If font.StrikeThrough <> CInt(WdConstants.wdUndefined) Then ff("strikeThrough") = (font.StrikeThrough = -1) Else ff("strikeThrough") = "mixed"
            If font.DoubleStrikeThrough <> CInt(WdConstants.wdUndefined) Then ff("doubleStrikeThrough") = (font.DoubleStrikeThrough = -1) Else ff("doubleStrikeThrough") = "mixed"
            If font.Subscript <> CInt(WdConstants.wdUndefined) Then ff("subscript") = (font.Subscript = -1) Else ff("subscript") = "mixed"
            If font.Superscript <> CInt(WdConstants.wdUndefined) Then ff("superscript") = (font.Superscript = -1) Else ff("superscript") = "mixed"

            If font.Color <> CType(WdConstants.wdUndefined, WdColor) Then
                ff("color") = font.Color.ToString()
                ff("colorRGB") = ColorToRGB(font.Color)
            Else
                ff("color") = "mixed"
            End If

            Try
                If rng.HighlightColorIndex <> CType(WdConstants.wdUndefined, WdColorIndex) Then
                    ff("highlightColor") = rng.HighlightColorIndex.ToString()
                Else
                    ff("highlightColor") = "mixed"
                End If
            Catch : ff("highlightColor") = "none" : End Try

            If font.AllCaps <> CInt(WdConstants.wdUndefined) Then ff("allCaps") = (font.AllCaps = -1) Else ff("allCaps") = "mixed"
            If font.SmallCaps <> CInt(WdConstants.wdUndefined) Then ff("smallCaps") = (font.SmallCaps = -1) Else ff("smallCaps") = "mixed"

            Try
                ff("scaling") = font.Scaling
                ff("spacing") = font.Spacing
                ff("position") = font.Position
                ff("kerning") = font.Kerning
            Catch : End Try

            Try
                Dim fontObj As Object = font
                ff("themeFont") = fontObj.ThemeFont.ToString()
            Catch : End Try

            Try
                Dim fontObj As Object = font
                ff("themeColor") = fontObj.ThemeColor.ToString()
                ff("themeTint") = fontObj.TintAndShade
            Catch : End Try

        Catch ex As Exception
            ff("error") = ex.Message
        End Try
        Return ff
    End Function

    Private Function ExtractListFormat(rng As Word.Range) As JObject
        Dim lf As New JObject()
        Try
            Dim listFormat As Word.ListFormat = rng.ListFormat

            If listFormat.ListType = WdListType.wdListNoNumbering Then
                lf("hasList") = False
                lf("listType") = "none"
            Else
                lf("hasList") = True
                lf("listType") = listFormat.ListType.ToString()
                lf("listLevelNumber") = listFormat.ListLevelNumber
                lf("listString") = listFormat.ListString

                Try
                    If listFormat.ListTemplate IsNot Nothing Then
                        Dim template As Word.ListTemplate = listFormat.ListTemplate
                        lf("listTemplateOutlineNumbered") = template.OutlineNumbered

                        Dim level As Word.ListLevel = template.ListLevels(listFormat.ListLevelNumber)
                        Dim levelInfo As New JObject()
                        levelInfo("numberFormat") = level.NumberFormat
                        levelInfo("numberStyle") = level.NumberStyle.ToString()
                        levelInfo("textPosition") = level.TextPosition
                        levelInfo("tabPosition") = level.TabPosition
                        levelInfo("numberPosition") = level.NumberPosition
                        levelInfo("alignment") = level.Alignment.ToString()
                        levelInfo("startAt") = level.StartAt
                        lf("currentLevelFormat") = levelInfo
                    End If
                Catch : End Try
            End If

        Catch ex As Exception
            lf("error") = ex.Message
        End Try
        Return lf
    End Function

    ''' <summary>
    ''' Extracts custom tab stop definitions.
    ''' </summary>
    Private Function ExtractTabStops(para As Word.Paragraph) As JArray
        Dim tabs As New JArray()
        Try
            For Each tabStop As Word.TabStop In para.TabStops
                Dim tab As New JObject()
                tab("position") = tabStop.Position
                tab("positionDescription") = "Position in points from left margin"
                tab("alignment") = tabStop.Alignment.ToString()
                tab("alignmentDescription") = "wdAlignTabLeft, wdAlignTabCenter, wdAlignTabRight, wdAlignTabDecimal, wdAlignTabBar"
                tab("leader") = tabStop.Leader.ToString()
                tab("leaderDescription") = "wdTabLeaderSpaces, wdTabLeaderDots, wdTabLeaderDashes, wdTabLeaderLines"
                tabs.Add(tab)
            Next
        Catch
        End Try
        Return tabs
    End Function

    ''' <summary>
    ''' Extracts paragraph border formatting.
    ''' </summary>
    Private Function ExtractBorders(para As Word.Paragraph) As JObject
        Dim borders As New JObject()
        Try
            Dim borderTypes As WdBorderType() = {
                WdBorderType.wdBorderTop,
                WdBorderType.wdBorderBottom,
                WdBorderType.wdBorderLeft,
                WdBorderType.wdBorderRight
            }
            Dim borderNames As String() = {"top", "bottom", "left", "right"}

            For i As Integer = 0 To borderTypes.Length - 1
                Try
                    Dim border As Word.Border = para.Borders(borderTypes(i))
                    If border.LineStyle <> WdLineStyle.wdLineStyleNone Then
                        Dim b As New JObject()
                        b("lineStyle") = border.LineStyle.ToString()
                        b("lineWidth") = border.LineWidth.ToString()
                        b("color") = border.Color.ToString()
                        b("colorRGB") = ColorToRGB(border.Color)
                        borders(borderNames(i)) = b
                    End If
                Catch
                End Try
            Next

            ' Box border distance
            Try
                borders("distanceFromText") = New JObject From {
                    {"top", para.Borders.DistanceFromTop},
                    {"bottom", para.Borders.DistanceFromBottom},
                    {"left", para.Borders.DistanceFromLeft},
                    {"right", para.Borders.DistanceFromRight}
                }
            Catch
            End Try

        Catch ex As Exception
            borders("error") = ex.Message
        End Try
        Return borders
    End Function

    ''' <summary>
    ''' Extracts paragraph shading/background properties.
    ''' </summary>
    Private Function ExtractShading(para As Word.Paragraph) As JObject
        Dim shading As New JObject()
        Try
            Dim s As Word.Shading = para.Shading
            shading("texture") = s.Texture.ToString()
            shading("textureDescription") = "wdTextureNone, wdTextureSolid, etc."

            If s.BackgroundPatternColor <> WdColor.wdColorAutomatic Then
                shading("backgroundColor") = s.BackgroundPatternColor.ToString()
                shading("backgroundColorRGB") = ColorToRGB(s.BackgroundPatternColor)
            End If

            If s.ForegroundPatternColor <> WdColor.wdColorAutomatic Then
                shading("foregroundColor") = s.ForegroundPatternColor.ToString()
                shading("foregroundColorRGB") = ColorToRGB(s.ForegroundPatternColor)
            End If

        Catch ex As Exception
            shading("error") = ex.Message
        End Try
        Return shading
    End Function


    ''' <summary>
    ''' Converts a Word color value to RGB hex string.
    ''' </summary>
    Private Function ColorToRGB(color As WdColor) As String
        Try
            Dim colorValue As Long = CLng(color)
            If colorValue < 0 Then Return "auto"

            Dim r As Integer = colorValue And &HFF
            Dim g As Integer = (colorValue >> 8) And &HFF
            Dim b As Integer = (colorValue >> 16) And &HFF

            Return $"#{r:X2}{g:X2}{b:X2}"
        Catch
            Return "unknown"
        End Try
    End Function

#End Region



#Region "DocStyle - Idempotent Style/List Helpers (PATCH)"

    Private Function ComputeListTemplateFingerprint(tpl As Word.ListTemplate) As String
        If tpl Is Nothing Then Return ""

        Try
            Dim sb As New StringBuilder()
            Dim maxLvl As Integer = 0
            Try
                maxLvl = Math.Min(9, tpl.ListLevels.Count)
            Catch
                maxLvl = 0
            End Try

            For lvl As Integer = 1 To maxLvl
                Try
                    Dim ll As Word.ListLevel = tpl.ListLevels(lvl)

                    Dim ns As String = ""
                    Try : ns = ll.NumberStyle.ToString() : Catch : ns = "" : End Try

                    Dim isBullet As Boolean = ns.IndexOf("bullet", StringComparison.OrdinalIgnoreCase) >= 0

                    If isBullet Then
                        ' Bullet formats are unstable in Word; fingerprint must include bullet font + char code.
                        Dim bFont As String = ""
                        Try
                            If ll.Font IsNot Nothing AndAlso Not String.IsNullOrEmpty(ll.Font.Name) Then
                                bFont = ll.Font.Name
                            End If
                        Catch
                            bFont = ""
                        End Try
                        If String.IsNullOrEmpty(bFont) Then bFont = "?"

                        Dim charCode As Integer = 0
                        Try
                            Dim nf As String = ll.NumberFormat
                            If Not String.IsNullOrEmpty(nf) Then
                                charCode = AscW(nf.Chars(0))
                            End If
                        Catch
                            charCode = 0
                        End Try

                        Dim numPos As Single = 0
                        Dim txtPos As Single = 0
                        Dim tabPos As Single = 0
                        Try : numPos = CSng(ll.NumberPosition) : Catch : End Try
                        Try : txtPos = CSng(ll.TextPosition) : Catch : End Try
                        Try : tabPos = CSng(ll.TabPosition) : Catch : End Try

                        sb.Append($"{lvl}:{ns}:BULLET:{bFont}:{charCode}:{Math.Round(numPos, 1)}:{Math.Round(txtPos, 1)}:{Math.Round(tabPos, 1)}|")
                    Else
                        Dim nf As String = ""
                        Try : nf = ll.NumberFormat : Catch : nf = "" : End Try

                        sb.Append($"{lvl}:{ns}:{nf}|")
                    End If
                Catch
                End Try
            Next

            Return sb.ToString()
        Catch
            Return ""
        End Try
    End Function
    Private Function TryGetStyleLinkedLevel(s As Word.Style, ByRef level As Integer) As Boolean
        level = 1
        Try
            level = s.ListLevelNumber
            If level <= 0 Then level = 1
            Return True
        Catch
            Return False
        End Try
    End Function



#End Region

End Class