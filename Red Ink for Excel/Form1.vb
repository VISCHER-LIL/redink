' Red Ink for Excel -- Chatbot Form Code
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See https://vischer.com/redink for more information.
'
' 18.11.2025
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Appache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core in unchanged form; Copyright (c) 2023 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes Nito.AsyncEx in unchanged form; Copyright (c) 2021 Stephen Cleary; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/StephenCleary/AsyncEx
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.


Imports System.Diagnostics
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.Globalization
Imports Microsoft.Office.Core
Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Reflection


' =============================================================================
' Excel Chatbot - Form1.vb — Reference overview
' =============================================================================
'
' Purpose
'   Chat UI for the Excel add-in ("Red Ink" / "Inky"). Provides:
'     - a lightweight chat window for calling the LLM via `SharedMethods.LLM`
'     - optional inclusion of worksheet content or selection in prompts
'     - model selection toggle (primary / secondary)
'     - simple persistence of transcript and window state via `My.Settings`
'     - parsing and execution of LLM-produced instructions via the host add-in
'
' High-level structure
'   - P/Invoke
'       - `SetForegroundWindow(hWnd As IntPtr)` — bring Excel window forward when executing commands
'
'   - Form-level fields & UI
'       - Buttons: `btnSend`, `btnCopy`, `btnCopyLastAnswer`, `btnClear`, `btnExit`, `btnSwitchModel`
'       - Checkboxes: `chkIncludeDocText` ("Include worksheet"), `chkIncludeselection`, `chkPermitCommands`, `chkStayOnTop`
'       - Panels: `pnlButtons`, `pnlCheckboxes`
'       - Text controls (designer): `txtChatHistory`, `txtUserInput`, `lblInstructions`
'       - State:
'           • `_context As ISharedContext` — SharedLibrary settings/context
'           • `_useSecondApi` — whether to call second model
'           • `_chatHistory` — List(Of (Role, Content)) holding conversation turns
'           • `OldChat`, `PreceedingNewline`, `UserLanguage`, `SystemPrompt`
'       - Constants / triggers: `ExtWSTrigger = "(addws)"`
'
' Lifecycle & initialization
'   - `New(context As ISharedContext)` — builds layout programmatically (TableLayoutPanel),
'       configures controls, stores `_context`.
'   - `frmAIChat_Load` — restores `My.Settings.LastChatHistory`, sets window title, icon, size,
'       wires event handlers, optionally shows a `WelcomeMessage`, warns on large worksheets.
'   - `frmAIChat_FormClosing` — persists `My.Settings.LastChatHistory`, `FormLocation`, `FormSize`.
'
' UI helpers
'   - `UpdateUIAsync(action As Action)` — marshals UI updates with `Invoke` when required.
'   - `AppendToChatHistory(text As String)` — append text to `txtChatHistory` thread-safely.
'   - `RemoveLastLineFromChatHistory()` — removes final line from transcript.
'   - Keyboard handlers:
'       • `UserInput_KeyDown` sends on Enter (Shift+Enter for newline)
'       • `oldUserInput_KeyDown` supports Ctrl+Enter
'       • `frmAIChat_KeyDown` closes on Escape (saves transcript within `_context.INI_ChatCap`)
'
' Conversation flow
'   - `btnSend_Click` — main handler:
'       1. Build `SystemPrompt` using `_context.SP_ChatExcel()` and checkbox flags.
'       2. Build conversation context via `BuildConversationString(_chatHistory)` and `OldChat`.
'       3. Validate Excel host availability when worksheet/selection inclusion is requested.
'       4. Optionally gather:
'           - entire worksheet (`Globals.ThisAddIn.ConvertRangeToString`)
'           - selection (`ConvertRangeToString` on the intersected selection)
'           - additional worksheets when user includes `ExtWSTrigger` via `Globals.ThisAddIn.GatherSelectedWorksheets()`
'       5. Construct `fullPrompt` with `<RANGEOFCELLS>` wrappers when passing range content.
'       6. Append user message to UI and `_chatHistory`.
'       7. Call `SharedMethods.LLM(_context, SystemPrompt, fullPrompt, ..., _useSecondApi, True)` asynchronously.
'       8. Sanitize LLM output:
'           - `RemoveMarkdownFormatting` usage (form is kept simple here)
'           - optionally extract `CommandsString` when `My.Settings.DoCommands` is true
'       9. Append assistant answer to UI and call `ExecuteAnyCommands(CommandsString)` when permitted.
'      10. Add assistant turn to `_chatHistory`.
'
' Welcome flow
'   - `WelcomeMessage()` — calls LLM for a localized greeting, appends result to transcript and `_chatHistory`.
'
' Conversation helpers
'   - `BuildConversationString(history)` — concatenates reversed history up to `_context.INI_ChatCap` characters.
'   - `GetCursorContext` is not present in this Excel form; selection context is gathered via `ConvertRangeToString`.
'
' Settings & model UI
'   - `btnSwitchModel_Click` — toggles `_useSecondApi` and updates the window title to reflect `INI_Model`/`INI_Model_2`.
'   - `UpdateDocumentCheckboxesState` (not present here) is in Word; Excel version disables model-dependent UI only when needed.
'   - `chkStayontop_Click`, `chkIncludeDocText_Click`, `chkIncludeselection_Click`, `chkPermitCommands_Click`
'       — manage `My.Settings` flags, validate selection via `IsSelectionEmpty(selection)`, and show warnings for large worksheets.
'
' Command parsing & execution
'   - `ParseCommands(input)` — parses command blocks of the form:
'       `[#command: @@argument1@@ §§argument2§§ #]`
'       • returns `List(Of ParsedCommand)` with `Command`, `Argument1`, `Argument2`
'       • regex-based parser tolerant to missing `arg2`
'   - `RemoveCommands(input)` — strips those command blocks from text and collapses excessive blank lines
'   - `ExecuteAnyCommands(commands As String)` — high-level executor:
'       1. Temporarily clear `TopMost` and bring Excel forward using `SetForegroundWindow`.
'       2. Calls `Globals.ThisAddIn.ParseLLMResponse(commands)` to obtain actionable `instructions` (list of `[Cell:...]` blocks).
'       3. If instructions exist:
'           - clears `Globals.ThisAddIn.undoStates`
'           - calls `Globals.ThisAddIn.ApplyLLMInstructions(instructions, True)` to apply changes (values, formulas, comments)
'           - updates undo UI via `Globals.Ribbons.Ribbon1.UpdateUndoButton()`
'       4. Restores form topmost and focus.
'
' Notes about command execution
'   - Actual cell-level changes, comments, and formula handling execute inside `ThisAddIn.ApplyLLMInstructions` and associated helpers (Excel ThisAddIn).
'   - Undo state is managed by `Globals.ThisAddIn.undoStates` so host ribbon UI can enable undo.
'   - `ExecuteAnyCommands` is deliberately simple: it transforms LLM result into the host add-in's instruction format (via `ParseLLMResponse`) and delegates application.
'
' Parsing utilities
'   - `ParsedCommand` helper DTO (properties: `Command`, `Argument1`, `Argument2`).
'   - `IsSelectionEmpty(selection As Excel.Range)` — checks intersection with `UsedRange` to detect a meaningful selection.
'
' Persistence & UX details
'   - Transcript persisted in `My.Settings.LastChatHistory` (capped by `_context.INI_ChatCap`).
'   - `My.Settings` stores checkbox preferences: `IncludeDocument`, `IncludeSelection`, `DoCommands`, `NotAlwaysOnTop`.
'   - `frmAIChat` uses `My.Resources.Red_Ink_Logo` as icon when available.
'   - Warns the user when including a large worksheet (uses `Globals.ThisAddIn.SizeOfWorksheet()` and `LargeWorksheetSize`).
'
' Threading & UI safety
'   - LLM calls are async/awaited; UI updates are marshaled via `UpdateUIAsync`.
'   - COM calls that read ranges are made synchronously on UI thread via `Globals.ThisAddIn` helpers.
'   - The form uses `Invoke` checks for thread-safe UI updates.
'
' Extension points & maintenance
'   - Add new chat commands: extend `ParseCommands` pattern and update callers that execute parsed commands (`ExecuteAnyCommands` or host `ApplyLLMInstructions`).
'   - For richer HTML chat like Word's version, a `WebBrowser` based renderer and Markdig pipeline could be reused (not present in this Excel form).
'   - When adding features that touch host ranges, reuse `Globals.ThisAddIn.ConvertRangeToString`, `GetFileContent`, and `ApplyLLMInstructions` to keep behavior consistent across hosts.
'   - Keep `_context` usage minimal in UI code; business logic should live in `SharedMethods` / `ThisAddIn`.
'
' Quick navigation (important methods)
'   - Constructor: `New(context As ISharedContext)`
'   - Load: `frmAIChat_Load`
'   - Send / LLM call: `btnSend_Click`
'   - Welcome: `WelcomeMessage`
'   - Command parsing: `ParseCommands`, `RemoveCommands`
'   - Command execution: `ExecuteAnyCommands`
'   - Helpers: `BuildConversationString`, `IsSelectionEmpty`, `AppendToChatHistory`, `UpdateUIAsync`
'
' =============================================================================


Public Class frmAIChat

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    Const AN As String = "Red Ink"
    Const AN5 As String = "Inky"   ' for Chatbox
    Private Const ExtWSTrigger As String = "(addws)"

    Private PreceedingNewline As String = ""
    Private OldChat As String = ""
    Private UserLanguage As String = New CultureInfo(Globals.ThisAddIn.Application.LanguageSettings.LanguageID(MsoAppLanguageID.msoLanguageIDUI)).Name
    Private SystemPrompt As String = ""


    Private WithEvents btnCopy As New System.Windows.Forms.Button() With {.Text = "Copy All", .AutoSize = True}
    Private WithEvents btnCopyLastAnswer As New System.Windows.Forms.Button() With {.Text = "Copy Last Answer", .AutoSize = True}
    Private WithEvents btnClear As New System.Windows.Forms.Button() With {.Text = "Clear", .AutoSize = True}
    Private WithEvents btnExit As New System.Windows.Forms.Button() With {.Text = "Close", .AutoSize = True}
    Private WithEvents btnSend As New System.Windows.Forms.Button() With {.Text = "Send", .AutoSize = True}
    Private WithEvents btnSwitchModel As New System.Windows.Forms.Button() With {.Text = "Switch Model", .AutoSize = True}
    Private WithEvents chkIncludeDocText As New System.Windows.Forms.CheckBox() With {.Text = "Include worksheet", .AutoSize = True, .Checked = My.Settings.IncludeDocument}
    Private WithEvents chkIncludeselection As New System.Windows.Forms.CheckBox() With {.Text = "Include selection", .AutoSize = True, .Checked = If(My.Settings.IncludeDocument, False, My.Settings.IncludeSelection)}
    Private WithEvents chkPermitCommands As New System.Windows.Forms.CheckBox() With {.Text = "Grant write access", .AutoSize = True, .Checked = My.Settings.DoCommands}
    Private WithEvents chkStayOnTop As New System.Windows.Forms.CheckBox() With {.Text = "Do not stay on top", .AutoSize = True, .Checked = My.Settings.NotAlwaysOnTop}


    Dim pnlButtons As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .Height = 40
    }

    Dim pnlCheckboxes As New FlowLayoutPanel() With {
        .Dock = DockStyle.Bottom,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .AutoSizeMode = AutoSizeMode.GrowAndShrink,
        .Height = 40
    }

    Private _context As ISharedContext = New SharedContext()

    ' Tracks whether we are using the second model/API.
    Private _useSecondApi As Boolean = False

    ' We keep the entire conversation in a List of (role, content).
    Private _chatHistory As New List(Of (Role As String, Content As String))


    Public Sub New(context As ISharedContext)
        ' This call is required by the designer.
        InitializeComponent()

        Me.AutoSize = False

        txtChatHistory.Multiline = True
        txtUserInput.Multiline = True

        ' 1) TableLayoutPanel anlegen
        Dim mainLayout As New TableLayoutPanel() With {
        .ColumnCount = 1,
        .RowCount = 5,
        .Dock = DockStyle.Fill,
        .AutoSize = False,
        .Padding = New Padding(10)   ' wird gleich überschrieben
    }

        ' 2) Spalten‑Breite auf 100 % setzen
        mainLayout.ColumnStyles.Clear()
        mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

        ' 3) Rechts 20 px Innenabstand
        mainLayout.Padding = New Padding(left:=10, top:=10, right:=20, bottom:=10)

        ' 4) Zeilen definieren
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        ' 5) Controls konfigurieren
        lblInstructions.AutoSize = True
        lblInstructions.Dock = DockStyle.Top
        txtChatHistory.Dock = DockStyle.Fill
        txtUserInput.Dock = DockStyle.Fill

        ' 6) Controls in die Tabelle packen
        mainLayout.Controls.Add(lblInstructions, 0, 0)
        mainLayout.Controls.Add(txtChatHistory, 0, 1)
        mainLayout.Controls.Add(txtUserInput, 0, 2)
        mainLayout.Controls.Add(pnlCheckboxes, 0, 3)
        mainLayout.Controls.Add(pnlButtons, 0, 4)

        ' 7) Form neu befüllen
        Me.Controls.Clear()
        Me.Controls.Add(mainLayout)

        _context = context
    End Sub


    ' Runs once when form loads.
    Private Async Sub frmAIChat_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.StartPosition = FormStartPosition.Manual
        Me.KeyPreview = True

        ' Restore saved chat text from My.Settings
        Dim previousChat As String = My.Settings.LastChatHistory
        If Not String.IsNullOrEmpty(previousChat) Then
            txtChatHistory.Text = previousChat
            OldChat = previousChat
            PreceedingNewline = Environment.NewLine
        End If

        ' Set the form's title and custom icon
        Me.Text = $"Chat (using " & If(_useSecondApi, _context.INI_Model_2, _context.INI_Model) & ")"
        Me.Font = New System.Drawing.Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.Sizable ' Ensure border supports icons
        Me.Icon = Icon.FromHandle(New Bitmap(My.Resources.Red_Ink_Logo).GetHicon())
        Me.TopMost = True ' Always on top

        ' Set the initial and minimum size of the form
        Me.MinimumSize = New Size(830, 521)

        If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
            Me.Location = My.Settings.FormLocation
            Me.Size = My.Settings.FormSize
        Else
            Me.StartPosition = FormStartPosition.CenterScreen
        End If

        AddHandler txtUserInput.KeyDown, AddressOf UserInput_KeyDown

        ' Set up instructions label
        lblInstructions.Text = $"Enter your question and click 'Send' or press Enter. Add '{ExtWSTrigger}' to pass along other open worksheets in your question. You can allow the chatbot to perform actions on your worksheet (change or comment cells): you can undo the last action, if needed."
        lblInstructions.AutoSize = True
        lblInstructions.Height = 50
        lblInstructions.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblInstructions.TextAlign = ContentAlignment.MiddleLeft

        ' FlowLayoutPanel for buttons

        pnlButtons.Padding = New Padding(0, 2, 8, 12)
        pnlButtons.Controls.Add(btnSend)
        pnlButtons.Controls.Add(btnCopyLastAnswer)
        pnlButtons.Controls.Add(btnCopy)
        pnlButtons.Controls.Add(btnClear)
        If _context.INI_SecondAPI Then pnlButtons.Controls.Add(btnSwitchModel)
        pnlButtons.Controls.Add(btnExit)

        pnlCheckboxes.Padding = New Padding(0, 1, 8, 1)
        pnlCheckboxes.Controls.Add(chkIncludeselection)
        pnlCheckboxes.Controls.Add(chkIncludeDocText)
        pnlCheckboxes.Controls.Add(chkPermitCommands)
        pnlCheckboxes.Controls.Add(chkStayOnTop)


        AddHandler btnCopy.Click, AddressOf btnCopy_Click
        AddHandler btnClear.Click, AddressOf btnClear_Click
        AddHandler btnSend.Click, AddressOf btnSend_Click
        AddHandler btnCopyLastAnswer.Click, AddressOf btnCopyLastAnswer_Click
        AddHandler btnSwitchModel.Click, AddressOf btnSwitchModel_Click
        AddHandler btnExit.Click, AddressOf btnExit_Click
        AddHandler chkIncludeselection.Click, AddressOf chkIncludeselection_Click
        AddHandler chkIncludeDocText.Click, AddressOf chkIncludeDocText_Click
        AddHandler chkPermitCommands.Click, AddressOf chkPermitCommands_Click
        AddHandler chkStayOnTop.Click, AddressOf chkStayontop_Click


        If String.IsNullOrWhiteSpace(txtChatHistory.Text) Then
            Dim result = Await WelcomeMessage()
        Else
            txtChatHistory.SelectionStart = txtChatHistory.Text.Length
            txtChatHistory.ScrollToCaret()

        End If

        If Globals.ThisAddIn.SizeOfWorksheet() > Globals.ThisAddIn.LargeWorksheetSize And chkIncludeDocText.Checked Then
            ShowCustomMessageBox($"Because this worksheet is large (a range of {Globals.ThisAddIn.SizeOfWorksheet()} cells, even if not all are used), it may slow down your interaction with the chatbot, because each time you send a question, the entire worksheet will be passed to {AN5}. If you want to speed up, include only a selection only.")
        End If

        If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()

    End Sub



    ' When the user clicks Send, we call the LLM with context.
    ' Then append the AI response to the conversation.

    Private Async Sub btnSend_Click(sender As Object, e As EventArgs)
        Dim userPrompt As String = txtUserInput.Text.Trim()
        If userPrompt = "" Then Return

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatExcel().Replace("{UserLanguage}", UserLanguage) & $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("MMMM dd, yyyy hh:mm tt")}. Only if you are expressly asked you can say that you have been developped by David Rosenthal of the law firm VISCHER in Switzerland. " & If(chkIncludeDocText.Checked, "\nYou have access to the user's document. \n", "") & If(chkIncludeselection.Checked, "\nYou have access to a selection of user's document. \n ", "") & If(My.Settings.DoCommands, _context.SP_Add_ChatExcel_Commands, _context.SP_Add_Chat_NoCommands)
            Dim conversationSoFar As String = BuildConversationString(_chatHistory)
            If Not String.IsNullOrWhiteSpace(OldChat) Then
                conversationSoFar += "\n" & OldChat
                OldChat = ""
            End If

            Dim appGuard As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
            If (chkIncludeDocText.Checked Or chkIncludeselection.Checked) AndAlso
                                (appGuard Is Nothing _
                               OrElse appGuard.Workbooks Is Nothing _
                               OrElse appGuard.Workbooks.Count = 0 _
                               OrElse appGuard.ActiveWorkbook Is Nothing _
                               OrElse appGuard.ActiveSheet Is Nothing) Then

                ShowCustomMessageBox("There is no active Excel worksheet. Please open or activate a workbook, then try again.")
                Return
            End If

            ' Optionally include Excel worksheet cells or selection
            Dim docText As String = ""
            Dim selectiontext As String = ""
            Dim selectedcells As String = ""
            Dim InsertWS As String = ""
            If chkIncludeDocText.Checked Then
                Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
                Dim selectedRange As Excel.Range = ws.UsedRange
                docText = Globals.ThisAddIn.ConvertRangeToString(selectedRange, True)
            End If
            If chkIncludeselection.Checked Or chkIncludeDocText.Checked Then
                Dim appx As Excel.Application = Globals.ThisAddIn.Application
                Dim selected As Excel.Range = appx.Selection
                Dim used As Excel.Range = appx.ActiveSheet.UsedRange
                Dim intersectedRange As Excel.Range = appx.Intersect(selected, used)
                If Not intersectedRange Is Nothing Then
                    If Not chkIncludeDocText.Checked Then
                        selectiontext = Globals.ThisAddIn.ConvertRangeToString(intersectedRange, True, True)
                        selectedcells = intersectedRange.Address(False, False)
                    Else
                        selectedcells = intersectedRange.Address(False, False)
                    End If
                End If
            End If

            If Not String.IsNullOrEmpty(userPrompt) And userPrompt.IndexOf(ExtWSTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                If Not chkIncludeDocText.Checked AndAlso Not chkIncludeselection.Checked Then
                    ShowCustomMessageBox("You cannot use the " & ExtWSTrigger & " trigger if you do not includ the worksheet or a selection of it - trigger ignored.")
                    InsertWS = ""
                Else
                    InsertWS = Globals.ThisAddIn.GatherSelectedWorksheets()
                    Debug.WriteLine($"GatherSelectedWorksheets returned: {Microsoft.VisualBasic.Left(InsertWS, 3000)}")
                    If String.IsNullOrWhiteSpace(InsertWS) Then
                        ShowCustomMessageBox("No content was found or an error occurred in gathering the additional worksheet(s) - doing without them.")
                        InsertWS = ""
                    ElseIf InsertWS.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then
                        ShowCustomMessageBox($"An error occured gathering the additional worksheet(s) ({InsertWS.Substring(6).Trim()}) - doing without them.")
                        InsertWS = ""
                    ElseIf InsertWS.StartsWith("NONE", StringComparison.OrdinalIgnoreCase) Then
                        ShowCustomMessageBox($"There are no other worksheets to add - doing without them.")
                        InsertWS = ""
                    End If

                End If
                userPrompt = Regex.Replace(userPrompt, Regex.Escape(ExtWSTrigger), "", RegexOptions.IgnoreCase)
            End If

            ' Construct the full prompt
            Dim fullPrompt As New StringBuilder()

            Dim app As Excel.Application = Globals.ThisAddIn.Application
            Dim workbookName As String = app.ActiveWorkbook.Name
            Dim worksheetName As String = app.ActiveSheet.Name
            Dim combinedName As String = workbookName & " - " & worksheetName

            If Not String.IsNullOrEmpty(docText) Then
                fullPrompt.AppendLine("You have access to the user's worksheet. The user's current worksheet is '" & combinedName & "' and has the following content: <RANGEOFCELLS>" & docText & "</RANGEOFCELLS>")
                If String.IsNullOrEmpty(selectiontext) Then
                    fullPrompt.AppendLine("The user has not selected any cells in this worksheet '" & combinedName & "'.")
                Else
                    fullPrompt.AppendLine("In the user's current worksheet '" & combinedName & "' the user has selected the following cells: " & selectedcells)
                End If
            ElseIf Not String.IsNullOrEmpty(selectiontext) Then
                fullPrompt.AppendLine("You have access to the user's worksheet. The user's current worksheet is '" & combinedName & "' and the user has selected the following cells: <RANGEOFCELLS>" & selectiontext & "</RANGEOFCELLS>")
            ElseIf chkIncludeselection.Checked Then
                fullPrompt.AppendLine("The user has granted you access to a selection of the worksheet '" & combinedName & "' but it is empty.")
            ElseIf chkIncludeDocText.Checked Then
                fullPrompt.AppendLine("The user has granted you access to the worksheet '" & combinedName & "' but the entire worksheet is empty.")
            End If
            If Not InsertWS.IsNullOrWhiteSpace(InsertWS) Then
                fullPrompt.AppendLine("The user also provided you access to the following additional worksheet(s): " & InsertWS)
            End If


            fullPrompt.AppendLine("User: " & userPrompt)
            fullPrompt.AppendLine("The conversation so far (not including any previously added worksheet content):\n" & conversationSoFar)

            ' Update UI on the UI thread
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(PreceedingNewline & "You: " & userPrompt.TrimEnd() & Environment.NewLine & Environment.NewLine)
                                    txtUserInput.Clear()
                                    PreceedingNewline = Environment.NewLine
                                End Sub)

            _chatHistory.Add(("user", userPrompt.TrimEnd()))

            ' Add a placeholder for AI response while waiting
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory($"{AN5}: Thinking...")
                                End Sub)


            ' Call the LLM function asynchronously
            Dim aiResponse As String = Await SharedMethods.LLM(_context, SystemPrompt, fullPrompt.ToString(), "", "", 0, _useSecondApi, True)
            aiResponse = aiResponse.TrimEnd()
            aiResponse = aiResponse.Replace($"{vbCrLf}* ", vbCrLf & ChrW(8226) & " ").Replace($"{vbCr}* ", vbCr & ChrW(8226) & " ").Replace($"{vbLf}* ", vbLf & ChrW(8226) & " ")
            aiResponse = aiResponse.Replace($"  *  ", "  " & ChrW(8226) & "  ")
            aiResponse = RemoveMarkdownFormatting(aiResponse)
            'aiResponse = aiResponse.Replace("**", "").Replace("_", "").Replace("`", "")

            Dim CommandsString As String = ""
            If My.Settings.DoCommands Then
                CommandsString = aiResponse
            End If

            Await UpdateUIAsync(Sub()
                                    RemoveLastLineFromChatHistory()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponse.TrimEnd().Replace(vbCrLf, Environment.NewLine).Replace(vbLf, Environment.NewLine) & Environment.NewLine)
                                    If My.Settings.DoCommands And Not String.IsNullOrWhiteSpace(CommandsString) Then
                                        ExecuteAnyCommands(CommandsString)
                                    End If
                                    txtUserInput.Text = ""
                                    If String.IsNullOrEmpty(txtUserInput.Text) Then txtUserInput.Focus()
                                End Sub)

            _chatHistory.Add(("assistant", aiResponse.TrimEnd()))

        Catch ex As System.Exception
            MsgBox("Error in btnSend_Click: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub



    Private Async Function WelcomeMessage() As Task(Of String)

        Try
            ' Build entire conversation so far into one string for context
            SystemPrompt = _context.SP_ChatExcel().Replace("{UserLanguage}", UserLanguage) & $" Your name is '{AN5}'. The current date and time is: {DateTime.Now.ToString("F")}. "
            txtUserInput.Text = ""

            ' Call the LLM function asynchronously
            Dim aiResponse As String = Await SharedMethods.LLM(_context, SystemPrompt, $"Welcome the user in {UserLanguage} by (1) referring to the time of day based on the current time in {UserLanguage} , such as in 'good morning', and (2) asking in {UserLanguage} what you can do, but do not say your name.", "", "", 0, _useSecondApi, True)

            aiResponse = aiResponse.Replace(vbLf, "").Replace(vbCr, "").Replace(vbCrLf, "") & vbCrLf
            aiResponse = aiResponse.Replace("**", "").Replace("_", "").Replace("`", "")

            ' Remove the "Thinking..." placeholder and update AI response on the UI thread
            Await UpdateUIAsync(Sub()
                                    AppendToChatHistory(Environment.NewLine & $"{AN5}: " & aiResponse)
                                End Sub)

            _chatHistory.Add(("assistant", aiResponse))

            PreceedingNewline = Environment.NewLine

            Return ""

        Catch ex As System.Exception
            'MsgBox("Error in WelcomeMessage: " & ex.Message, MsgBoxStyle.Critical)
            Return ""
        End Try
    End Function


    ' Helper method to ensure UI updates occur on the correct thread
    Private Async Function UpdateUIAsync(action As System.Action) As System.Threading.Tasks.Task
        If InvokeRequired Then
            Await System.Threading.Tasks.Task.Run(Sub() Me.Invoke(action))
        Else
            action()
        End If
    End Function

    Private Sub AppendToChatHistory(text As String)
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() txtChatHistory.AppendText(text))
        Else
            txtChatHistory.AppendText(text)
        End If
    End Sub



    Private Sub RemoveLastLineFromChatHistory()
        If txtChatHistory.InvokeRequired Then
            txtChatHistory.Invoke(Sub() RemoveLastLineFromChatHistory())
        Else
            Dim lines As String() = txtChatHistory.Lines
            If lines.Length > 0 Then
                txtChatHistory.Lines = lines.Take(lines.Length - 1).ToArray()
            End If
        End If
    End Sub

    Private Sub chkStayontop_Click(sender As Object, e As EventArgs)
        Me.TopMost = Not Me.TopMost
        My.Settings.NotAlwaysOnTop = Me.TopMost
        My.Settings.Save()
    End Sub

    Private Sub chkPermitCommands_Click(sender As Object, e As EventArgs)
        My.Settings.DoCommands = Not My.Settings.DoCommands

        My.Settings.Save()
    End Sub


    Private Sub chkIncludeselection_Click(sender As Object, e As EventArgs) Handles chkIncludeselection.Click
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim selection As Excel.Range = TryCast(app.Selection, Excel.Range)

        ' Check if selection is valid and contains data
        If selection Is Nothing OrElse IsSelectionEmpty(selection) Then
            chkIncludeselection.Checked = False
        ElseIf chkIncludeDocText.Checked Then
            chkIncludeDocText.Checked = False
        End If

        My.Settings.IncludeSelection = chkIncludeselection.Checked
        My.Settings.Save()

    End Sub

    Private Function IsSelectionEmpty(selection As Excel.Range) As Boolean
        Dim ws As Excel.Worksheet = selection.Worksheet
        Dim app As Excel.Application = ws.Application

        ' build the range of all cells that "mean something"
        Dim infoRange As Excel.Range = ws.UsedRange

        ' see if any of those intersect the user’s selection
        Dim intersected As Excel.Range = Nothing
        Try
            intersected = app.Intersect(selection, infoRange)
        Catch ex As System.Exception
            ' should never really happen, but just in case
            Return True
        End Try

        ' if nothing in common, it's empty
        Return (intersected Is Nothing) OrElse (intersected.Cells.Count = 0)
    End Function



    Private Sub chkIncludeDocText_Click(sender As Object, e As EventArgs)

        If chkIncludeselection.Checked Then
            chkIncludeselection.Checked = False
        End If
        My.Settings.IncludeDocument = chkIncludeDocText.Checked
        My.Settings.Save()

        If Globals.ThisAddIn.SizeOfWorksheet() > Globals.ThisAddIn.LargeWorksheetSize And chkIncludeDocText.Checked Then
            ShowCustomMessageBox($"Because this worksheet is large (a range of {Globals.ThisAddIn.SizeOfWorksheet()} cells, even if not all are used), it may slow down your interaction with the chatbot, because each time you send a question, the entire worksheet will be passed to {AN5}. If you want to speed up, include only a selection only.")
        End If

    End Sub


    ' Copies the entire conversation to the clipboard.

    Private Sub btnCopy_Click(sender As Object, e As EventArgs)
        My.Computer.Clipboard.SetText(txtChatHistory.Text)
    End Sub


    ' Copies only the last AI answer to the clipboard.

    Private Sub btnCopyLastAnswer_Click(sender As Object, e As EventArgs)
        Dim lastAssistantMsg = _chatHistory.Where(Function(x) x.Role = "assistant").LastOrDefault()
        If lastAssistantMsg.Content IsNot Nothing Then
            My.Computer.Clipboard.SetText(lastAssistantMsg.Content)
        Else
            SharedMethods.ShowCustomMessageBox("No last AI answer available.")
        End If
    End Sub


    ' Switches the model from model1 to model2 and vice versa.

    Private Sub btnSwitchModel_Click(sender As Object, e As EventArgs)
        _useSecondApi = Not _useSecondApi
        Me.Text = $"Chat (using " & If(_useSecondApi, _context.INI_Model_2, _context.INI_Model) & ")"
    End Sub


    ' Clears the conversation from both the UI and saved settings.

    Private Sub btnClear_Click(sender As Object, e As EventArgs)

        _chatHistory.Clear()
        txtChatHistory.Clear()
        OldChat = ""
        PreceedingNewline = ""
        My.Settings.LastChatHistory = ""
        My.Settings.Save()
        Dim result = WelcomeMessage()
    End Sub


    ' Press Escape to close. Also button-based exit.

    Private Sub frmAIChat_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Dim conversation As String = txtChatHistory.Text
            If conversation.Length > _context.INI_ChatCap Then
                conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
            End If
            My.Settings.LastChatHistory = conversation
            My.Settings.Save()
            Close()
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs)
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation
        My.Settings.Save()
        Close()
    End Sub

    Private Sub frmAIChat_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Save the chat history before the form closes
        Dim conversation As String = txtChatHistory.Text
        If conversation.Length > _context.INI_ChatCap Then
            conversation = conversation.Substring(conversation.Length - _context.INI_ChatCap)
        End If
        My.Settings.LastChatHistory = conversation

        ' Save the form's location and size to My.Settings
        If Me.WindowState = FormWindowState.Normal Then
            My.Settings.FormLocation = Me.Location
            My.Settings.FormSize = Me.Size
        Else
            ' If the form is minimized or maximized, save the restored bounds
            My.Settings.FormLocation = Me.RestoreBounds.Location
            My.Settings.FormSize = Me.RestoreBounds.Size
        End If
        My.Settings.Save()

    End Sub


    ' Trigger the Send button on Ctrl+Enter in the user input textbox.

    Private Sub oldUserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.Enter Then
            btnSend.PerformClick()
            e.Handled = True
        End If
    End Sub

    ' Trigger the Send button on Enter, allow Shift+Enter for new line
    Private Sub UserInput_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If e.Shift Then
                ' Allow Shift+Enter to insert a new line (default behavior)
                Return
            Else
                ' Enter alone sends the message
                e.SuppressKeyPress = True
                btnSend.PerformClick()
                e.Handled = True
            End If
        End If
    End Sub


    ' Builds the conversation history as a single string.

    Private Function BuildConversationString(history As List(Of (Role As String, Content As String))) As String
        Dim sb As New StringBuilder()
        Dim totalLength As Integer = 0
        Dim maxLength As Integer = _context.INI_ChatCap

        ' Iterate through the history in reverse order (most recent messages first)
        For Each msg In history.AsEnumerable().Reverse()
            Dim message As String
            If msg.Role = "user" Then
                message = $"User: {msg.Content}{Environment.NewLine}"
            Else
                message = $"{AN5}: {msg.Content}{Environment.NewLine}"
            End If

            ' Check if adding this message will exceed the limit
            If totalLength + message.Length > maxLength Then
                ' If so, truncate the message to fit within the limit
                Dim remainingLength As Integer = maxLength - totalLength
                If remainingLength > 0 Then
                    sb.Insert(0, message.Substring(0, remainingLength))
                End If
                Exit For
            Else
                ' Otherwise, append the full message
                sb.Insert(0, message)
                totalLength += message.Length
            End If
        Next

        Return sb.ToString()
    End Function


    Private Sub pnlCheckboxes_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Public Class ParsedCommand
        Public Property Command As String
        Public Property Argument1 As String
        Public Property Argument2 As String
    End Class


    ' Parses the input string for embedded commands of the format:
    ' [#command: @@argument1@@ §§argument2§§ #]
    ' Returns a List of ParsedCommand objects.
    ' argument2 is optional; if not present, it defaults to "".
    Private Function ParseCommands(input As String) As List(Of ParsedCommand)
        Dim results As New List(Of ParsedCommand)

        Try
            ' Regex Explanation:
            ' \[#       matches literal [#
            ' (?<cmd>[^:]+)    matches 1 or more characters that are not :, captured as group "cmd"
            ' :\s*     matches a colon and optional whitespace
            ' @@(?<arg1>[^@]+)@@   matches @@ + 1 or more non-@ chars + @@, capturing as group "arg1"
            ' \s*      optional whitespace
            ' (?:§§(?<arg2>[^§]+)§§)?  optional group: §§ + 1 or more non-§ chars + §§, captured as "arg2"
            ' \s*      optional whitespace
            ' #\]      literal #]
            ' The "?" after the group means "optional"

            Dim pattern As String = "\[#(?<cmd>[^:]+):\s*@@(?<arg1>[^@]+)@@\s*(?:§§(?<arg2>[^§]*)§§)?\s*#\]"
            Dim regex As New Regex(pattern)

            For Each m As Match In regex.Matches(input)
                Dim pc As New ParsedCommand()

                pc.Command = m.Groups("cmd").Value.Trim()
                pc.Argument1 = m.Groups("arg1").Value.Trim()

                ' If arg2 wasn't found, it might be blank
                If m.Groups("arg2") IsNot Nothing Then
                    pc.Argument2 = m.Groups("arg2").Value.Trim().Replace("\r\n", vbCrLf).Replace("\n", vbCrLf).Replace("\r", vbCrLf)
                End If

                If String.IsNullOrEmpty(pc.Argument2) Then
                    pc.Argument2 = ""
                Else
                    pc.Argument1 = pc.Argument1.Replace("\r\n", ".*").Replace("\n", ".*").Replace("\r", ".*")
                    pc.Argument1 = pc.Argument1.Replace(vbCrLf, ".*").Replace(vbCr, ".*").Replace(vbLf, ".*")
                End If

                If Not results.Any(Function(x) x.Command = pc.Command AndAlso x.Argument1 = pc.Argument1 AndAlso x.Argument2 = pc.Argument2) Then
                    results.Add(pc)
                End If
            Next

        Catch ex As System.Exception
            MsgBox("Error in ParseCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return results
    End Function


    ' Removes all commands of the format:
    ' [#command: @@argument1@@ §§argument2§§ #]
    ' from the input string.
    Public Function RemoveCommands(input As String) As String
        Dim output As String = input
        Try
            ' Remove the commands along with immediate surrounding whitespace and line breaks
            Dim commandPattern As String = "\s*[\r\n]*\s*\[#[^:]+:\s*@@[^@]+@@\s*(?:§§[^§]*§§)?\s*#\]\s*[\r\n]*\s*"
            Dim regex As New Regex(commandPattern)
            output = regex.Replace(input, "")

            ' Collapse multiple consecutive line breaks into a single line break
            Dim whitespacePattern As String = "[\r\n]{3,}"
            Dim collapseRegex As New Regex(whitespacePattern)
            output = collapseRegex.Replace(output, Environment.NewLine)

        Catch ex As System.Exception
            MsgBox("Error in RemoveCommands: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return output
    End Function


    Private CommandsList As String = ""

    Public Sub ExecuteAnyCommands(commands As String)
        Dim topmost As Boolean = Me.TopMost
        Me.TopMost = False

        Dim instructions As New List(Of String)
        instructions = Globals.ThisAddIn.ParseLLMResponse(commands)

        If instructions.Count > 0 Then
            ' Bring Excel window to front (instead of Application.Activate())
            Dim hwnd As IntPtr = CType(Globals.ThisAddIn.Application.Hwnd, IntPtr)
            SetForegroundWindow(hwnd)

            System.Threading.Thread.Sleep(200)

            Globals.ThisAddIn.undoStates.Clear()
            Globals.ThisAddIn.ApplyLLMInstructions(instructions, True)
            Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()

        End If


        Me.TopMost = topmost
        Me.Focus()
    End Sub


End Class
