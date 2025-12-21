' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

' =============================================================================
' File: ThisAddIn.Commands.vb
' Purpose: Defines command methods invoked by the UI for translation, correction,
'          improvement, anonymization, shortening, party switching, freestyle prompt
'          execution, chat display, help window, and settings management. Methods
'          gather user input, set state variables, and delegate processing to
'          ProcessSelectedRange or other helper routines.
'
' Architecture:
'   - Each public async command:
'       1. Calls Application.DoEvents.
'       2. Optionally acquires the current Excel.Range selection.
'       3. Gathers user input via ShowCustomInputBox / prompt library.
'       4. Sets global/context variables (TranslateLanguage, Context, OldParty/NewParty, etc.).
'       5. Invokes ProcessSelectedRange with a system prompt constant (SP_*), passing flags
'          controlling range vs cell-by-cell vs formula handling, color check, pane usage,
'          file object inclusion, batch path, worksheet content, and shortening percentage.
'   - Freestyle(...) parses prefix triggers (CellByCellPrefix, CellByCellPrefix2, TextPrefix,
'       TextPrefix2, BubblesPrefix, PanePrefix, BatchPrefix, PurePrefix, ExtTrigger,
'       ExtWSTrigger, ColorTrigger, ObjectTrigger, ObjectTrigger2) to derive execution mode
'       and augment OtherPrompt with file content, worksheet text, color analysis, file object
'       reference, clipboard object, or batch directory processing.
'   - Batch processing: collects a target insertion line (LineNumber) and directory path
'       (BatchPath), validating presence of allowedExtensions.
'   - File/worksheet inclusion: replaces ExtTrigger occurrences with tagged file content; adds
'       additional worksheet text via GatherSelectedWorksheets when ExtWSTrigger is present.
'   - Settings: builds dictionaries of setting names/descriptions and displays a settings window;
'       refreshes menus via AddContextMenu using a SplashScreen.
'   - ShowChatForm: creates and positions frmAIChat with persisted size/location (My.Settings).
'   - HelpMeInky: lazy-instantiates HelpMeInky and calls ShowRaised.
'   - Return values: Methods declare a Boolean result variable from ProcessSelectedRange or other
'       operations; result is not explicitly returned (no Return statement) in this file.
' =============================================================================

Imports System.Diagnostics
Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Translates the selected range to INI_Language1 using SP_Translate. Sets TranslateLanguage.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function InLanguage1() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = INI_Language1
        Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
    End Function

    ''' <summary>
    ''' Translates the selected range to INI_Language2 using SP_Translate. Sets TranslateLanguage.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function InLanguage2() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = INI_Language2
        Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
    End Function

    ''' <summary>
    ''' Prompts for a target language, selects current range if available, translates using SP_Translate.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned. Exits early if no language provided.</returns>
    Public Async Function InOther() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            If selectedRange IsNot Nothing Then
                selectedRange.Select()
            End If
            Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
        End If
    End Function

    ''' <summary>
    ''' Prompts for a target language and translates including formulas (third flag True) using SP_Translate.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function InOtherFormulas() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, True, False, True, False)
        End If
    End Function

    ''' <summary>
    ''' Performs correction on selected range using SP_Correct.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function Correct() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await ProcessSelectedRange(SP_Correct, True, False, False, False, True, False)
    End Function

    ''' <summary>
    ''' Prompts for optional context, defaults to "n/a" if blank, then improves writing using SP_WriteNeatly.
    ''' Requires a selected range.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Returns False early if no selection. Result variable not explicitly returned.</returns>
    Public Async Function Improve() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If
        Context = Trim(SLib.ShowCustomInputBox("Please provide the context that should be taken into account, if any:", $"{AN} Write Neatly", True))
        If String.IsNullOrWhiteSpace(Context) Then
            Context = "n/a"
        End If
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If
        Dim result As Boolean = Await ProcessSelectedRange(SP_WriteNeatly, True, False, False, False, True, False)
    End Function

    ''' <summary>
    ''' Anonymizes selected range content using SP_Anonymize.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function Anonymize() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await ProcessSelectedRange(SP_Anonymize, True, False, False, False, True, False)
    End Function

    ''' <summary>
    ''' Prompts for a shortening percentage (1–99), computes average/max length of unprotected non-formula cells,
    ''' then shortens content using SP_Shorten with percentage parameter.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Returns False early on invalid selection or input. Result variable not explicitly returned.</returns>
    Public Async Function Shorten() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If
        Dim totalLength As Integer = 0
        Dim maxLength As Integer = 0
        Dim cellCount As Integer = 0
        For Each cell As Excel.Range In selectedRange.Cells
            If Not CellProtected(cell) AndAlso Not CBool(cell.HasFormula) Then
                Dim cellText As String = CStr(cell.Value)
                If Not String.IsNullOrEmpty(cellText) Then
                    Dim textLength As Integer = getnumberofwords(cellText)
                    totalLength += textLength
                    If textLength > maxLength Then
                        maxLength = textLength
                    End If
                    cellCount += 1
                End If
            End If
        Next
        Dim averageLength As Double = If(cellCount > 0, totalLength / cellCount, 0)
        Dim UserInput As String
        Dim ShortenPercentValue As Integer = 0
        Do
            UserInput = Trim(SLib.ShowCustomInputBox($"Enter the percentage by which the text of each selected cell should be shortened (the cells have have of average {averageLength:n1} words and {maxLength} at max; {ShortenPercent}% will cut approx. " & (averageLength * ShortenPercent / 100) & " words in average):", $"{AN} Shortener", True, CStr(ShortenPercent) & "%"))
            If String.IsNullOrEmpty(UserInput) Then
                Return False
            End If
            UserInput = UserInput.Replace("%", "").Trim()
            If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid percentage between 1 and 99.")
            End If
        Loop
        If ShortenPercentValue = 0 Then Return False
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If
        Dim result As Boolean = Await ProcessSelectedRange(SP_Shorten, True, False, False, False, True, False, ShortenPercentValue, False)
    End Function

    ''' <summary>
    ''' Prompts for old and new party names separated by a comma, sets OldParty/NewParty, then switches occurrences using SP_SwitchParty.
    ''' Requires a selected range.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Returns False early if selection missing or input invalid. Result variable not explicitly returned.</returns>
    Public Async Function SwitchParty() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If
        Dim UserInput As String
        Do
            UserInput = Trim(SLib.ShowCustomInputBox("Please provide the original party name and the new party name, separated by a comma (example: Elvis Presley, Taylor Swift):", $"{AN} Switch Party", True))
            If String.IsNullOrEmpty(UserInput) Then
                Return False
            End If
            Dim parts() As String = UserInput.Split(","c)
            If parts.Length = 2 Then
                OldParty = parts(0).Trim()
                NewParty = parts(1).Trim()
                Exit Do
            Else
                ShowCustomMessageBox("Please enter two names separated by a comma.")
            End If
        Loop
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If
        Dim result As Boolean = Await ProcessSelectedRange(SP_SwitchParty, True, False, False, False, True, False)
    End Function

    ''' <summary>
    ''' Shows or creates the chat form (frmAIChat) and restores size/location from My.Settings; brings form to front.
    ''' </summary>
    Public Sub ShowChatForm()
        If chatForm Is Nothing OrElse chatForm.IsDisposed Then
            chatForm = New frmAIChat(_context)
            ' Set the location and size before showing the form
            If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
                chatForm.StartPosition = FormStartPosition.Manual
                chatForm.Location = My.Settings.FormLocation
                chatForm.Size = My.Settings.FormSize
            Else
                ' Default to center screen if no settings are available
                chatForm.StartPosition = FormStartPosition.Manual
                Dim screenBounds As System.Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea
                chatForm.Location = New System.Drawing.Point((screenBounds.Width - chatForm.Width) \ 2, (screenBounds.Height - chatForm.Height) \ 2)
                chatForm.Size = New Size(650, 500) ' Set default size if needed
            End If
        End If
        ' Show and bring the form to the front
        chatForm.Show()
        chatForm.BringToFront()
    End Sub

    ''' <summary>
    ''' Freestyle execution without alternate model (UseSecondAPI = False). Delegates to Freestyle(False).
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned.</returns>
    Public Async Function FreestyleNM() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await Freestyle(False)
    End Function

    ''' <summary>
    ''' Freestyle execution with optional alternate model selection (if INI_AlternateModelPath provided).
    ''' Restores defaults if originalConfigLoaded is True after execution.
    ''' </summary>
    ''' <returns>Task(Of Boolean). Result variable not explicitly returned. Exits early if model selection canceled.</returns>
    Public Async Function FreestyleAM() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
            If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                originalConfigLoaded = False
                Exit Function
            End If
        End If
        Dim result As Boolean = Await Freestyle(True)
        If originalConfigLoaded Then
            RestoreDefaults(_context, originalConfig)
            originalConfigLoaded = False
        End If
    End Function

    ''' <summary>
    ''' Central freestyle method parsing prefixes/triggers to set execution flags (cell-by-cell, text-only, comments, pane, batch, color check,
    ''' file object, worksheet inclusion). Handles prompt library, stored last prompt, file content insertion, worksheet gathering,
    ''' batch directory selection, line number input, and delegates to ProcessSelectedRange with appropriate system prompt or raw prompt.
    ''' </summary>
    ''' <param name="UseSecondAPI">True to use second API/model path; affects object trigger availability.</param>
    ''' <returns>Task(Of Boolean). Result variables not explicitly returned. Returns False early in multiple validation failure cases.</returns>
    Public Async Function Freestyle(ByVal UseSecondAPI As Boolean) As Task(Of Boolean)
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        Dim NoSelectedCells As Boolean = False
        Dim DoClipboard As Boolean = False
        Dim DoFormulas As Boolean = True
        Dim DoBubbles As Boolean = False
        Dim DoColor As Boolean = False
        Dim DoFileObject As Boolean = False
        Dim DoFileObjectClip As Boolean = False
        Dim DoPane As Boolean = False
        Dim DoBatch As Boolean = False
        Dim BatchPath As String = ""
        Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
        Dim PureInstruct As String = $"; use '{PurePrefix}' for direct prompting"
        Dim DefaultPrefix As String = INI_DefaultPrefix
        Dim DefaultPrefixText As String = ""
        If selectedRange Is Nothing Then
            NoSelectedCells = True
        End If
        Dim DoRange As Boolean = True
        Dim CBCInstruct As String = $"with '{CellByCellPrefix}' or '{CellByCellPrefix2} if the instruction should be executed cell-by-cell"
        Dim TextInstruct As String = $"use '{TextPrefix}' or '{TextPrefix2}' if the instruction should apply cell-by-cell, but only to text cells"
        Dim BatchInstruct As String = $"use '{BatchPrefix}' if to process a directory of files"
        Dim BubblesInstruct As String = $"use '{BubblesPrefix}' for inserting comments only"
        Dim PaneInstruct As String = $"use '{PanePrefix}' for using the pane"
        Dim ExtInstruct As String = $"; insert '{ExtTrigger}' (multiple times) for text of (a) file(s) (txt, docx, pdf) or '{ExtWSTrigger}' to add more worksheet(s)"
        Dim AddonInstruct As String = $"; add '{ColorTrigger}' to check for colorcodes"
        Dim ObjectInstruct As String = $"; add '{ObjectTrigger}'/'{ObjectTrigger2}' for adding a file object"
        Dim FileObject As String = ""
        Dim InsertWS As String = ""
        If UseSecondAPI Then
            If Not String.IsNullOrWhiteSpace(INI_APICall_Object_2) Then
                AddonInstruct += ObjectInstruct.Replace("; add", ",")
                DoFileObject = True
            End If
        Else
            If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                AddonInstruct += ObjectInstruct.Replace("; add", ",")
                DoFileObject = True
            End If
        End If
        Dim PromptLibInstruct As String = ""
        If INI_PromptLib Then
            PromptLibInstruct = " or press 'OK' for the prompt library"
        End If
        If DefaultPrefix.Trim() <> "" Then
            DefaultPrefixText = $" (default prefix: '{DefaultPrefix}')"
        End If
        Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix)
        }
        If Not NoSelectedCells Then
            OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected cells (start {CBCInstruct}; {TextInstruct}; {PaneInstruct}; {BatchInstruct}; {BubblesInstruct})" & PromptLibInstruct & PureInstruct & ExtInstruct & AddonInstruct & LastPromptInstruct & DefaultPrefixText & ":", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons))
        Else
            OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute {PromptLibInstruct} (the result will be shown to you before inserting anything into your worksheet); {PaneInstruct}{BatchInstruct}{PureInstruct}{ExtInstruct}{AddonInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons))
            DoRange = True
        End If
        If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then
            Dim promptlibresult As (String, Boolean, Boolean, Boolean)
            promptlibresult = ShowPromptSelector(INI_PromptLibPath, INI_PromptLibPathLocal, Nothing, Nothing)
            OtherPrompt = promptlibresult.Item1
            DoClipboard = promptlibresult.Item4
            If OtherPrompt = "" Then
                Return False
            End If
        Else
            If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return False
        End If
        ' Check if OtherPrompt starts with a word ending with a colon
        If Not String.IsNullOrWhiteSpace(OtherPrompt) Then
            Dim firstWord As String = OtherPrompt.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
            If firstWord IsNot Nothing AndAlso Not firstWord.EndsWith(":"c) Then
                Dim prefix As String = DefaultPrefix.Trim()
                ' Ensure prefix ends with colon and space
                If prefix <> "" AndAlso Not prefix.EndsWith(":"c) Then
                    prefix &= ":"
                End If
                OtherPrompt = prefix & " " & OtherPrompt.Trim()
                OtherPrompt = OtherPrompt.Trim()
            End If
        End If
        My.Settings.LastPrompt = OtherPrompt
        My.Settings.Save()
        If Not SharedMethods.ProcessParameterPlaceholders(OtherPrompt) Then
            ShowCustomMessageBox("Freestyle canceled.", $"{AN} Freestyle")
            Return False
        End If
        If OtherPrompt.StartsWith(CellByCellPrefix, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(CellByCellPrefix.Length).Trim()
            DoRange = False
        End If
        If OtherPrompt.StartsWith(CellByCellPrefix2, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(CellByCellPrefix2.Length).Trim()
            DoRange = False
        End If
        If OtherPrompt.StartsWith(TextPrefix, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(TextPrefix.Length).Trim()
            DoRange = False
            DoFormulas = False
        End If
        If OtherPrompt.StartsWith(TextPrefix2, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(TextPrefix2.Length).Trim()
            DoRange = False
            DoFormulas = False
        End If
        If OtherPrompt.StartsWith(BubblesPrefix, StringComparison.OrdinalIgnoreCase) And selectedRange IsNot Nothing Then
            OtherPrompt = OtherPrompt.Substring(BubblesPrefix.Length).Trim()
            DoBubbles = True
            DoRange = True
        End If
        If OtherPrompt.StartsWith(PanePrefix, StringComparison.OrdinalIgnoreCase) And DoRange Then
            OtherPrompt = OtherPrompt.Substring(PanePrefix.Length).Trim()
            DoPane = True
            DoRange = True
        End If
        If OtherPrompt.StartsWith(BatchPrefix, StringComparison.OrdinalIgnoreCase) Then
            OtherPrompt = OtherPrompt.Substring(BatchPrefix.Length).Trim()
            DoPane = False
            DoRange = True
            DoBatch = True
            Try
                ' --- Excel context (as requested) ---
                Dim activeCell As Microsoft.Office.Interop.Excel.Range = Application.ActiveCell
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = CType(Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                Dim currentRow As System.Int32 = 1
                If activeCell IsNot Nothing Then
                    currentRow = System.Convert.ToInt32(activeCell.Row, System.Globalization.CultureInfo.InvariantCulture)
                End If
                Dim maxRow As System.Int32 = System.Convert.ToInt32(ws.Rows.Count, System.Globalization.CultureInfo.InvariantCulture)
                ' --- Loop until user provides a valid line number or cancels/ESC ---
                Dim lineNumberAnswer As System.String = System.String.Empty
                Do
                    lineNumberAnswer = ShowCustomInputBox(
                        "Please provide the starting line number at which the results should be inserted (1 for first line, 2 for second line etc.):", $"{AN} Freestyle Batch",
                        True,
                        currentRow.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    )
                    If lineNumberAnswer Is Nothing Then
                        Return Nothing
                    End If
                    lineNumberAnswer = lineNumberAnswer.Trim()
                    If lineNumberAnswer.Length = 0 Then
                        Return Nothing
                    End If
                    If System.String.Compare(lineNumberAnswer, "ESC", System.StringComparison.OrdinalIgnoreCase) = 0 Then
                        Return Nothing
                    End If
                    Dim parsedRow As System.Int32
                    If Not System.Int32.TryParse(lineNumberAnswer, parsedRow) OrElse parsedRow < 1 OrElse parsedRow > maxRow Then
                        ShowCustomMessageBox("The provided value is not a valid line number between 1 and " &
                                             maxRow.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
                    Else
                        LineNumber = parsedRow
                        Exit Do
                    End If
                Loop
                ' --- Directory picker (browse or type) ---
                Dim initialPath As System.String
                If (Application.ActiveWorkbook IsNot Nothing) AndAlso
                   (Application.ActiveWorkbook.Path IsNot Nothing) AndAlso
                   (Application.ActiveWorkbook.Path.Length > 0) Then
                    initialPath = Application.ActiveWorkbook.Path
                Else
                    initialPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments)
                End If
                Using dlg As New System.Windows.Forms.FolderBrowserDialog()
                    dlg.Description = "Select the directory that contains the batch text files."
                    dlg.ShowNewFolderButton = False
                    dlg.SelectedPath = initialPath
                    Dim result As System.Windows.Forms.DialogResult = dlg.ShowDialog()
                    If result <> System.Windows.Forms.DialogResult.OK Then
                        Return Nothing
                    End If
                    Dim selectedPath As System.String = dlg.SelectedPath
                    If System.String.IsNullOrWhiteSpace(selectedPath) OrElse Not System.IO.Directory.Exists(selectedPath) Then
                        ShowCustomMessageBox("No directory was selected or it does not exist.")
                        Return Nothing
                    End If
                    Dim hasAny As System.Boolean = False
                    ' Enumerate files and check extensions
                    For Each filePath As System.String In System.IO.Directory.EnumerateFiles(selectedPath, "*.*", System.IO.SearchOption.TopDirectoryOnly)
                        Dim ext As System.String = System.IO.Path.GetExtension(filePath)
                        If allowedExtensions.Contains(ext) Then
                            hasAny = True
                            Exit For
                        End If
                    Next
                    ' Handle case when no allowed files exist
                    If Not hasAny Then
                        ShowCustomMessageBox(
                            "The selected directory does not contain any files of the expected types: " &
                            System.String.Join(", ", allowedExtensions) & "."
                        )
                        Return Nothing
                    End If
                    BatchPath = selectedPath
                End Using
            Catch ex As System.Exception
                ShowCustomMessageBox("GetLineNumber and Path resulted in an Error: " & ex.Message)
                Return Nothing
            End Try
        End If
        If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            OtherPrompt = OtherPrompt.Replace(ObjectTrigger, "(a file object follows)").Trim()
        ElseIf DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
            OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a file object follows)").Trim()
            DoFileObjectClip = True
        Else
            DoFileObject = False
        End If
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If
        If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(ColorTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            DoColor = True
            OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ColorTrigger), "", RegexOptions.IgnoreCase)
        End If
        If Not String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            ' Count total occurrences first (case-insensitive) so inserted file text containing {doc} does not trigger extra loops.
            Dim totalOccurrences As Integer = Regex.Matches(OtherPrompt, Regex.Escape(ExtTrigger), RegexOptions.IgnoreCase).Count
            ' Detect if a placeholder occurrence is already enclosed by any tag, e.g. <tag>...{doc}...</tag>
            Dim wrappedPattern As String =
                "<(?<name>[A-Za-z][\w\-]*)\b[^>]*>[^<]*" & Regex.Escape(ExtTrigger) & "[^<]*</\k<name>>"
            If totalOccurrences = 1 Then
                ' Single-occurrence behavior with optional auto-wrapping
                DragDropFormLabel = ""
                DragDropFormFilter = ""
                Dim doc As String = Await GetFileContent(Nothing, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object), True)
                If String.IsNullOrWhiteSpace(doc) Then
                    ShowCustomMessageBox("The file you have selected is empty or not supported - exiting.")
                    Return False
                End If
                Dim isWrapped As Boolean = Regex.IsMatch(OtherPrompt, wrappedPattern, RegexOptions.IgnoreCase)
                Dim replacementText As String = If(isWrapped, doc, $"<document>{doc}</document>")
                OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ExtTrigger), replacementText, RegexOptions.IgnoreCase)
                ShowCustomMessageBox($"This file will be included in your prompt where you have referred to {ExtTrigger}: " & vbCrLf & vbCrLf & doc)
            Else
                ' Multi-occurrence behavior: prompt separately for each placeholder
                For occurrence As Integer = 1 To totalOccurrences
                    Dim idx As Integer = OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase)
                    If idx < 0 Then Exit For
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    Dim docPart As String = Await GetFileContent(Nothing, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object), True)
                    If String.IsNullOrWhiteSpace(docPart) Then
                        Dim answer As Integer = ShowCustomYesNoBox($"The file you selected for occurrence #{occurrence} is empty, not supported or you cancelled the upload. Do you want to continue or abort?", "Continue", "Abort")
                        If answer = 2 Then Return False
                    End If

                    Dim replacementText As String = ""

                    If Not String.IsNullOrEmpty(docPart) Then
                        ' Determine if this specific occurrence is already wrapped by any tag pair
                        Dim isWrappedThis As Boolean = False
                        Dim mcol As MatchCollection = Regex.Matches(OtherPrompt, wrappedPattern, RegexOptions.IgnoreCase)
                        For Each m As Match In mcol
                            If idx >= m.Index AndAlso idx < m.Index + m.Length Then
                                isWrappedThis = True
                                Exit For
                            End If
                        Next
                        replacementText = If(isWrappedThis, docPart, $"<document{occurrence}>{docPart}</document{occurrence}>")

                    End If

                    ' Replace only the first remaining occurrence (manual replacement keeps later placeholders intact)
                    OtherPrompt = OtherPrompt.Substring(0, idx) & replacementText & OtherPrompt.Substring(idx + ExtTrigger.Length)

                    If Not String.IsNullOrWhiteSpace(docPart) Then
                        ShowCustomMessageBox($"This file will be included at occurrence #{occurrence} (of {totalOccurrences}) where you used {ExtTrigger}:" &
                                         vbCrLf & vbCrLf & docPart)
                    End If
                Next
            End If
        End If
        If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(ExtWSTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            If Not DoRange Then
                ShowCustomMessageBox($"{ExtWSTrigger} cannot be combined with cell by cell processing - exiting.")
                Return False
            End If
            InsertWS = GatherSelectedWorksheets(True)
            Debug.WriteLine($"GatherSelectedWorksheets returned: {Left(InsertWS, 3000)}")
            If String.IsNullOrWhiteSpace(InsertWS) Then
                ShowCustomMessageBox("No content was found or an error occurred in gathering the additional worksheet(s) - exiting.")
                Return False
            ElseIf InsertWS.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"An error occured gathering the additional worksheet(s) ({InsertWS.Substring(6).Trim()}) - exiting.")
                Return False
            ElseIf InsertWS.StartsWith("NONE", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"There are no other worksheets to add - exiting.")
                Return False
            End If
            OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ExtWSTrigger), "", RegexOptions.IgnoreCase)
        End If
        If DoFileObject Then
            If DoFileObjectClip Then
                FileObject = "clipboard"
            Else
                DragDropFormLabel = "All file types that are supported by your LLM."
                DragDropFormFilter = "Supported Files|*.*"
                FileObject = GetFileName()
                DragDropFormLabel = ""
                DragDropFormFilter = ""
                If String.IsNullOrWhiteSpace(FileObject) Then
                    ShowCustomMessageBox("No file object has been selected - will abort. You can try again (use Ctrl-P to re-insert your prompt).")
                    Return False
                End If
            End If
        End If
        If OtherPrompt.StartsWith(PurePrefix, StringComparison.OrdinalIgnoreCase) Then
            OtherPrompt = OtherPrompt.Substring(PurePrefix.Length).Trim()
            Dim result As Boolean = Await ProcessSelectedRange(OtherPrompt, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS)
        Else
            If Not NoSelectedCells Then
                If DoRange Then
                    Dim result As Boolean = Await ProcessSelectedRange(SP_RangeOfCells, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS, BatchPath)
                Else
                    Dim result As Boolean = Await ProcessSelectedRange(SP_FreestyleText, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS)
                End If
            Else
                Dim result As Boolean = Await ProcessSelectedRange(SP_RangeOfCells, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS, BatchPath)
            End If
        End If
    End Function

    Private _win As HelpMeInky = Nothing

    ''' <summary>
    ''' Lazy-instantiates HelpMeInky window (HelpMeInky) and displays it using ShowRaised.
    ''' </summary>
    Public Sub HelpMeInky()
        If _win Is Nothing OrElse _win.IsDisposed Then
            _win = New HelpMeInky(_context, RDV)
        End If
        ' No owner needed
        _win.ShowRaised()
    End Sub

    ''' <summary>
    ''' Builds settings name/description dictionaries, shows settings window, then updates menus via AddContextMenu wrapped in a SplashScreen.
    ''' </summary>
    Public Sub ShowSettings()
        Dim Settings As New Dictionary(Of String, String) From {
            {"Temperature", "Temperature of {model}"},
            {"Timeout", "Timeout of {model}"},
            {"Temperature_2", "Temperature of {model2}"},
            {"Timeout_2", "Timeout of {model2}"},
            {"DoubleS", "Convert '" & ChrW(223) & "' to 'ss'"},
            {"NoEmDash", "Convert em to en dash"},
            {"PreCorrection", "Additional instruction for prompts"},
            {"PostCorrection", "Prompt to apply after queries"},
            {"Language1", "Default translation language 1"},
            {"Language2", "Default translation language 2"},
            {"PromptLibPath", "Prompt library file"},
            {"PromptLibPathLocal", "Prompt library file (local)"},
            {"DefaultPrefix", "Default prefix to use in 'Freestyle'"}
        }
        Dim SettingsTips As New Dictionary(Of String, String) From {
            {"Temperature", "The higher, the more creative the LLM will be (0.0-2.0)"},
            {"Timeout", "In milliseconds"},
            {"Temperature_2", "The higher, the more creative the LLM will be (0.0-2.0)"},
            {"Timeout_2", "In milliseconds"},
            {"DoubleS", "For Switzerland"},
            {"NoEmDash", "This will convert long dashes typically generated by LLMs but that are not commonly used (thus suggesting that the text has been AI generated)"},
            {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
            {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
            {"Language1", "The language (in English) that will be used for the first quick access button in the ribbon"},
            {"Language2", "The language (in English) that will be used for the second quick access button in the ribbon"},
            {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"},
            {"PromptLibPathLocal", "The filename (including path, support environmental variables) for your local prompt library (if any)"},
            {"DefaultPrefix", "You can define here the default prefix to use within 'Freestyle' if no other prefix is used (will be added authomatically)."}
        }
        ShowSettingsWindow(Settings, SettingsTips)
        Dim splash As New SplashScreen("Updating menu following your changes ...")
        splash.Show()
        splash.Refresh()
        AddContextMenu()
        splash.Close()
    End Sub

End Class