' Part of: Red Ink for Word
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Public Shared LastFreestyleModelConfig As ModelConfig
    Public Shared LastFreestyleWasAM As Boolean = False
    Public Shared LastFreestylePrompt As String = ""

    Public Async Sub FreeStyleNM()
        If INILoadFail() Then Return
        FreeStyle(False)

        My.Settings.LastFreestyleModelConfig = Nothing
        My.Settings.LastFreestyleWasAM = False
        My.Settings.LastFreestylePrompt = My.Settings.LastPrompt
        My.Settings.Save()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

    End Sub


    Public Async Sub FreeStyleAM()
        If INILoadFail() Then Return

        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then

            If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                originalConfigLoaded = False
                Return
            End If

        End If

        LastFreestyleModelConfig = GetCurrentConfig(_context)

        FreeStyle(True)

        My.Settings.LastFreestyleModelConfig = LastFreestyleModelConfig
        My.Settings.LastFreestyleWasAM = True
        My.Settings.LastFreestylePrompt = My.Settings.LastPrompt
        My.Settings.Save()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

    End Sub

    Public Async Sub FreeStyleRepeat()
        If INILoadFail() Then Return

        Dim LastFreestylePrompt As String = My.Settings.LastFreestylePrompt

        originalConfig = GetCurrentConfig(_context)
        originalConfigLoaded = False

        If String.IsNullOrWhiteSpace(LastFreestylePrompt) Then
            ShowCustomMessageBox("No last Freestyle command has been stored.")
            Return
        End If

        If My.Settings.LastFreestyleWasAM Then
            LastFreestyleModelConfig = My.Settings.LastFreestyleModelConfig

            If LastFreestyleModelConfig IsNot Nothing Then
                Dim ErrorFlag As Boolean = True
                ApplyModelConfig(_context, LastFreestyleModelConfig, ErrorFlag)
                If ErrorFlag Then
                    ShowCustomMessageBox("There was an error assigning the last model configuration. Aborting.")
                    Return
                End If
                originalConfigLoaded = True
            End If
        End If

        FreeStyle(My.Settings.LastFreestyleWasAM, My.Settings.LastFreestylePrompt)

    End Sub






    Public Async Sub FreeStyle(UseSecondAPI As Boolean, Optional LastPrompt As String = "")
        If INILoadFail() Then Return
        Try
            OtherPrompt = ""
            SysPrompt = ""
            InsertDocs = ""
            MyStyleInsert = ""

            Dim NoText As Boolean = False
            Dim DoMarkup As Boolean = False
            Dim DoClipboard As Boolean = False
            Dim DoBubbles As Boolean = False
            Dim DoInplace As Boolean = Override(INI_ReplaceText2, INI_ReplaceText2Override)
            Dim MarkupMethod As Integer = Override(INI_MarkupMethodWord, INI_MarkupMethodWordOverride)
            Dim DoLib As Boolean = False
            Dim DoNet As Boolean = False
            Dim DoTPMarkup As Boolean = False
            Dim TPMarkupName As String = ""
            Dim KeepFormatCap = INI_KeepFormatCap
            Dim DoKeepFormat As Boolean = INI_KeepFormat2
            Dim DoKeepParaFormat As Boolean = INI_KeepParaFormatInline
            Dim DoFileObject As Boolean = False
            Dim DoFileObjectClip As Boolean = False
            Dim DoPane As Boolean = False
            Dim DoNewDoc As Boolean = False
            Dim DoChunks As Boolean = False
            Dim ChunkSize As Integer = 1
            Dim NoFormatAndFieldSaving As Boolean = False
            Dim DoSlides As Boolean = False
            Dim DoMyStyle As Boolean = False
            Dim DoMultiModel As Boolean = True
            Dim DoBubblesExtract As Boolean = False
            Dim DoPushback As Boolean = False

            Dim MarkupInstruct As String = $"start With '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"with '{InPlacePrefix}'/'{AddPrefix} for replacing/adding to the selection"
            Dim BubblesInstruct As String = $"with '{BubblesPrefix}' for having your text commented"
            Dim PushbackInstruct As String = $"with '{PushbackPrefix}' for responding to comments only"
            Dim SlidesInstruct As String = $"with '{SlidesPrefix}' for adding to a Powerpoint file"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}', '{NewdocPrefix}' or '{PanePrefix}' for separate output"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim ExtInstruct As String = $"; inlcude '{ExtTrigger}' (multiple times) for text of (a) file(s) (txt, docx, pdf) or '{AddDocTrigger}' for an open Word doc"
            Dim TPMarkupInstruct As String = $"; add '{TPMarkupTriggerInstruct}' if revisions [of user] should be pointed out to the LLM"
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}/{SameAsReplaceTrigger}' for overriding formatting defaults"
            Dim AllInstruct As String = $"; add '{AllTrigger}' to select all"
            Dim MyStyleInstruct As String = $"; add '{MyStyleTrigger}' to apply your personal style"
            Dim LibInstruct As String = $"; add '{LibTrigger}' for library search"
            Dim NetInstruct As String = $"; add '{NetTrigger}' for internet search"
            Dim PureInstruct As String = $"; use '{PurePrefix}' for direct prompting"
            Dim ChunkInstruct As String = $"; add '{ChunkTrigger}' for iterating through the text"
            Dim BubblesExtractInstruct As String = $"; add '{BubblesExtractTrigger}' for including bubble comments"
            Dim ObjectInstruct As String = $"; add '{ObjectTrigger}'/'{ObjectTrigger2}' for adding a file object"
            Dim MultiModelInstruct As String = $"; add '{MultiModelTrigger}' for multiple models"
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
            Dim FileObject As String = ""
            Dim SlideDeck As String = ""

            Dim DefaultPrefix As String = INI_DefaultPrefix
            Dim DefaultPrefixText As String = ""

            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection

            If selection.Type = WdSelectionType.wdSelectionIP Then NoText = True

            Dim AddOnInstruct As String = AllInstruct

            If Not NoText Then
                AddOnInstruct += NoFormatInstruct.Replace("; add", ", ")
                AddOnInstruct += TPMarkupInstruct.Replace("; add", ", ")
                AddOnInstruct += ChunkInstruct.Replace("; add", ", ")
                AddOnInstruct += BubblesExtractInstruct.Replace("; add", ", ")
            End If
            If INI_Lib Then
                AddOnInstruct += LibInstruct.Replace("; add", ",")
            End If
            If INI_ISearch Then
                AddOnInstruct += NetInstruct.Replace("; add", ", ")
            End If
            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) Then
                AddOnInstruct += MyStyleInstruct.Replace("; add", ", ")
            End If
            If UseSecondAPI Then
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object_2) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    AddOnInstruct += MultiModelInstruct.Replace("; add", ", ")
                End If
            Else
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
            End If

            Dim lastCommaIndex As Integer = AddOnInstruct.LastIndexOf(","c)
            If lastCommaIndex <> -1 Then
                AddOnInstruct = AddOnInstruct.Substring(0, lastCommaIndex) & ", and" & AddOnInstruct.Substring(lastCommaIndex + 1)
            End If

            If DefaultPrefix.Trim() <> "" Then
                DefaultPrefixText = $" (default prefix: '{DefaultPrefix}')"
            End If


            If LastPrompt.Trim() = "" Then
                If Not NoText Then
                    Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix),
                            System.Tuple.Create("OK, do a markup", $"Use this to automatically insert '{MarkupPrefixDiff}' as a prefix.", MarkupPrefixDiff)
                        }

                    OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {ClipboardInstruct}, {InplaceInstruct}, {BubblesInstruct}, {PushbackInstruct} or {SlidesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons).Trim()
                Else
                    Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix)
                        }
                    OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct} or {SlidesInstruct}){PromptLibInstruct}{ExtInstruct}{AddOnInstruct}{PureInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons).Trim()
                End If
            Else
                OtherPrompt = LastPrompt
            End If

            Debug.WriteLine($"OtherPrompt: '{OtherPrompt}'")

            SelectedText = ""

            If Not NoText Then

                SelectedText = selection.Text

                If String.Equals(OtherPrompt.Trim(), "codebasis", StringComparison.OrdinalIgnoreCase) Then
                    SLib.WriteToRegistry(RemoveCR(RegPath_CodeBasis), RemoveCR(selection.Text))
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Return
                End If
                If OtherPrompt.StartsWith("inipath", StringComparison.OrdinalIgnoreCase) Then
                    SLib.WriteToRegistry(RemoveCR(RegPath_IniPath), RemoveCR(selection.Text))
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    Return
                End If
                If String.Equals(OtherPrompt.Trim(), "encode", StringComparison.OrdinalIgnoreCase) Then
                    Dim Key As String = CodeAPIKey(RemoveCR(selection.Text))
                    SLib.PutInClipboard(Key)
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    selection.TypeText(vbCrLf & "Encoded key (also in clipboard):" & vbCrLf & Key)
                    selection.ParagraphFormat.Hyphenation = CInt(False) ' Turn off hyphenation
                    SLib.PutInClipboard(Key)
                    Return
                End If

                If String.Equals(OtherPrompt.Trim(), "decode", StringComparison.OrdinalIgnoreCase) Then
                    Dim Key As String = DeCodeAPIKey(RemoveCR(selection.Text))
                    SLib.PutInClipboard(Key)
                    selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                    selection.TypeText(vbCrLf & "Decoded key (also in clipboard):" & vbCrLf & Key)
                    selection.ParagraphFormat.Hyphenation = CInt(False) ' Turn off hyphenation
                    Return
                End If

                If OtherPrompt.StartsWith("convertmarkdown", StringComparison.OrdinalIgnoreCase) Then
                    Dim trailingCR = (SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbLf) Or SelectedText.EndsWith(vbCr))
                    InsertTextWithMarkdown(selection, SelectedText, trailingCR, True)
                    Return
                End If

            End If
            If String.Equals(OtherPrompt.Trim(), "domain", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"{AN} is running in the domain '{GetDomain()}' and configured to run in {If(String.IsNullOrEmpty(SLib.alloweddomains), "any domain ('alloweddomains' has not been set).", "'" & SLib.alloweddomains & "'.")}", "")
                Return
            End If
            If String.Equals(OtherPrompt.Trim(), "model", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox("I am using the " & INI_Model & " model as my primary model with a default timeout of " & (INI_Timeout / 1000) & " seconds (" & Microsoft.VisualBasic.Strings.Format(INI_Timeout / 60000, "0.00") & " minutes)." & If(INI_MaxOutputToken > 0, "The maximum output token length is " & INI_MaxOutputToken & ".", ""))
                Return
            End If
            If String.Equals(OtherPrompt.Trim(), "terms", StringComparison.OrdinalIgnoreCase) Then
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selection.TypeText(vbCrLf & If(INI_UsageRestrictions = "", "No usage restrictions or permissions have been defined in the configuration file.", "The defined usage restrictions or permissions defined in the configuration file are: " & INI_UsageRestrictions) & vbCrLf)
                Return
            End If
            If String.Equals(OtherPrompt.Trim(), "anonymize", StringComparison.OrdinalIgnoreCase) Then
                AnonymizeSelection()
                Return
            End If
            If OtherPrompt.StartsWith("insertclipboard", StringComparison.OrdinalIgnoreCase) OrElse OtherPrompt.StartsWith("insertclip", StringComparison.OrdinalIgnoreCase) Then
                Call InsertClipboard()
                Return
            End If

            If OtherPrompt.StartsWith("generateresponsekey", StringComparison.OrdinalIgnoreCase) Or OtherPrompt.StartsWith("generateresponsetemplate", StringComparison.OrdinalIgnoreCase) Then

                If NoText Then
                    ShowCustomMessageBox("No text has been selected. Select the text containing both the JSON payload to interpret and what you want the output to look like (by referencing to the JSON fields and structure in natural text).")
                    Return
                End If

                Dim response As String = Await LLM(SP_GenerateResponseKey & vbCrLf & Code_JsonTemplateFormatter, vbCrLf & SelectedText, "", "", 0, UseSecondAPI)

                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                selection.InsertAfter(vbCrLf & vbCrLf & response)

                Return
            End If

            If OtherPrompt.StartsWith("editmystyle", StringComparison.OrdinalIgnoreCase) Then
                SLib.ShowTextFileEditor(ExpandEnvironmentVariables(INI_MyStylePath), "Edit your MyStyle prompt file (use 'Define MyStyle' to create new prompts automatically):")
                Return
            End If


            If OtherPrompt.StartsWith("definemystyle", StringComparison.OrdinalIgnoreCase) Then
                DefineMyStyle()
                Return
            End If

            If OtherPrompt.StartsWith("promptlog", StringComparison.OrdinalIgnoreCase) Then
                ShowAndEditPromptLog()
                Return
            End If

            If OtherPrompt.StartsWith("webagentcreator", StringComparison.OrdinalIgnoreCase) Then
                CreateModifyWebAgentScript()
                Return
            End If


            If String.Equals(OtherPrompt.Trim(), "webagent", StringComparison.OrdinalIgnoreCase) Then
                WebAgent()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "findhiddenprompts", StringComparison.OrdinalIgnoreCase) Then
                FindHiddenPrompts()
                Return
            End If


            If OtherPrompt.StartsWith("redinktest", StringComparison.OrdinalIgnoreCase) Then

                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                Dim filePath As String = System.IO.Path.Combine(desktopPath, "redinktest.txt")
                If File.Exists(filePath) Then
                    Dim testtextorig As String = File.ReadAllText(filePath).Replace("\n", vbCrLf)
                    Dim testtext As String = SLib.ShowCustomWindow("Testfile content:", testtextorig, "", AN, False, True, True, True)
                    If testtext <> "" And testtext <> "Pane" Then
                        If testtext = "Markdown" Then
                            Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                            Globals.ThisAddIn.Application.Selection.TypeParagraph()
                            InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & testtextorig, False)
                            Dim patternx As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                            If Regex.IsMatch(testtextorig, patternx) Then
                                Dim rng As Range = wordApp.Selection.Range
                                RestoreSpecialTextElements(rng)
                                rng.Document.Fields.Update()
                            End If
                        Else
                            SLib.PutInClipboard(testtext)
                        End If
                    ElseIf testtext = "Pane" Then
                        SP_MergePrompt_Cached = SP_MergePrompt
                        ShowPaneAsync(
                                                                        "Test Pane",
                                                                        testtextorig,
                                                                        "",
                                                                        AN,
                                                                        noRTF:=False,
                                                                        insertMarkdown:=True
                                                                        )
                    End If
                    Return
                Else
                    Return
                End If
            End If


            If String.Equals(OtherPrompt.Trim(), "switch", StringComparison.OrdinalIgnoreCase) Then
                selection.Range.Collapse(Direction:=Word.WdCollapseDirection.wdCollapseEnd)
                If INI_SecondAPI Then
                    SwitchModels(_context)
                    ShowCustomMessageBox("You have temporarily switched the two configured models. Primary is now '" & INI_Model & "', and secondary is '" & INI_Model_2 & "'.")
                Else
                    ShowCustomMessageBox("You have defined only one model ('" & INI_Model & "').")
                End If
                Return
            End If
            If String.Equals(OtherPrompt.Trim(), "version", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox("You are using " & Version & $" of {AN}. (c) by David Rosenthal, VISCHER. Go to https://vischer.com/{AN2} for more information. This copy of {AN} is set to expire on {LicensedTill.ToString("dd-MMM-yyyy")}", AN)
                Return
            End If
            If String.Equals(OtherPrompt.Trim(), "reset", StringComparison.OrdinalIgnoreCase) Then
                If ShowCustomYesNoBox($"Do you really want to reset your local configuration file and settings (if any) by removing non-mandatory entries? The current configuration file '{AN2}.ini' will NOT be saved to a '.bak' file. If you only want to reload the configuration settings for giving up any temporary changes, use 'reload' instead.", "Yes", "No") = 1 Then
                    INIloaded = False
                    ResetLocalAppConfig(_context)
                    MenusAdded = False
                    AddContextMenu()
                    ShowCustomMessageBox($"Following the reset, the configuration file '{AN2}.ini' has been be reloaded.")
                End If
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "speech", StringComparison.OrdinalIgnoreCase) Then
                Transcriptor()
                Return

            End If

            If OtherPrompt.StartsWith("readlocal", StringComparison.OrdinalIgnoreCase) Then
                SpeakSelectedText()
                Return

            End If

            If OtherPrompt.StartsWith("clearlastprompt", StringComparison.OrdinalIgnoreCase) Then
                My.Settings.LastPrompt = ""
                My.Settings.LastFreestylePrompt = ""
                My.Settings.LastFreestyleModelConfig = Nothing
                My.Settings.LastFreestyleWasAM = False
                My.Settings.Save()
                Dim resultx = Globals.Ribbons.Ribbon1.InitializeAppAsync()
                ShowCustomMessageBox($"The last Freestyle prompt has been cleared.")

                Return

            End If

            If OtherPrompt.StartsWith("voiceslocal", StringComparison.OrdinalIgnoreCase) Then
                SelectVoiceByNumber()
                Return
            End If

            If OtherPrompt.StartsWith("voices2", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New TTSSelectionForm("Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", True) ' TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", True)
                    If frm.ShowDialog() = DialogResult.OK Then
                        ' Retrieve selected voices
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim outputPath As String = frm.SelectedOutputPath
                        ' Use the selected values
                        If selectedVoices.Count > 0 Then
                            MessageBox.Show("Selected Voice(s): " & String.Join(", ", selectedVoices))
                        Else
                            MessageBox.Show("No voices selected.")
                        End If

                        If outputPath = "" Then
                            MessageBox.Show("Temporary output selected.")
                        Else
                            MessageBox.Show("Output path: " & outputPath)
                        End If
                    Else
                        MessageBox.Show("Voice selection was cancelled.")
                    End If
                End Using

                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "voices", StringComparison.OrdinalIgnoreCase) Then
                Using frm As New TTSSelectionForm("Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", False) ' TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voices you wish to use.", $"{AN} Text-to-Speech - Select Voices", False)
                    If frm.ShowDialog() = DialogResult.OK Then
                        ' Retrieve selected voices
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim outputPath As String = frm.SelectedOutputPath
                        ' Use the selected values
                        If selectedVoices.Count > 0 Then
                            MessageBox.Show("Selected Voice(s): " & String.Join(", ", selectedVoices))
                        Else
                            MessageBox.Show("No voices selected.")
                        End If

                        If outputPath = "" Then
                            MessageBox.Show("Temporary output selected.")
                        Else
                            MessageBox.Show("Output path: " & outputPath)
                        End If
                    Else
                        MessageBox.Show("Voice selection was cancelled.")
                    End If
                End Using

                Return
            End If

            If OtherPrompt.StartsWith("doccheck", StringComparison.OrdinalIgnoreCase) Then
                RunDocCheck()
                Return
            End If

            If OtherPrompt.StartsWith("findclause", StringComparison.OrdinalIgnoreCase) Then
                FindClause()
                Return
            End If

            If OtherPrompt.StartsWith("addclause", StringComparison.OrdinalIgnoreCase) Then
                AddClause()
                Return
            End If


            If OtherPrompt.StartsWith("createpodcast", StringComparison.OrdinalIgnoreCase) Then
                CreatePodcast()
                Return
            End If

            If OtherPrompt.StartsWith("readpodcast", StringComparison.OrdinalIgnoreCase) Then
                ReadPodcast(selection.Text)
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "read", StringComparison.OrdinalIgnoreCase) Then
                CreateAudio()
                Return
            End If

            If OtherPrompt.StartsWith("cleanmenu", StringComparison.OrdinalIgnoreCase) Then
                RemoveOldContextMenu()
                RemoveVeryOldContextMenu()
                MenusAdded = False
                AddContextMenu()
                Return
            End If

            If String.Equals(OtherPrompt.Trim(), "reload", StringComparison.OrdinalIgnoreCase) Then
                INIloaded = False
                InitializeConfig(False, True)
                MenusAdded = False
                AddContextMenu()
                ShowCustomMessageBox($"The configuration file '{AN2}.ini' has been be reloaded.")
                Return
            End If
            If String.Equals(OtherPrompt.Trim(), "settings", StringComparison.OrdinalIgnoreCase) Then
                ShowSettings()
                Return
            End If

            If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then

                Dim promptlibresult As (String, Boolean, Boolean, Boolean)

                promptlibresult = ShowPromptSelector(INI_PromptLibPath, INI_PromptLibPathLocal, Not NoText, Not NoText)

                OtherPrompt = promptlibresult.Item1
                DoMarkup = promptlibresult.Item2
                DoBubbles = promptlibresult.Item3
                DoClipboard = promptlibresult.Item4

                If OtherPrompt = "" Then
                    Return
                End If
            Else
                If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return
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
                Exit Sub
            End If

            If OtherPrompt.IndexOf(AllTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(AllTrigger, "").Trim()
                Dim document As Word.Document = application.ActiveDocument
                document.Content.Select()
                NoText = False
            End If

            If OtherPrompt.IndexOf(LibTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(LibTrigger, "").Trim()
                DoLib = True
            End If

            If OtherPrompt.IndexOf(TPMarkupTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(TPMarkupTrigger, "").Trim()
                DoTPMarkup = True
            End If

            If OtherPrompt.IndexOf(ChunkTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ChunkTrigger, "").Trim()
                DoChunks = True
            End If

            If OtherPrompt.IndexOf(BubblesExtractTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(BubblesExtractTrigger, "").Trim()
                DoBubblesExtract = True
                If DoChunks Then
                    ShowCustomMessageBox($"The '{BubblesExtractTrigger}' option cannot be used together with '{ChunkTrigger}' - the bubble comments will not be extracted.")
                    DoBubblesExtract = False
                End If
            End If

            ' Formatting Trigger

            If OtherPrompt.IndexOf(NoFormatTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(NoFormatTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger2, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(KFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger2, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger, "").Trim()
                DoKeepParaFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger2, "").Trim()
                DoKeepParaFormat = True
            End If
            If Not DoInplace Then
                If OtherPrompt.IndexOf(SameAsReplaceTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    OtherPrompt = OtherPrompt.Replace(SameAsReplaceTrigger, "").Trim()
                Else
                    NoFormatAndFieldSaving = True
                End If
            End If

            If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger, "(a file object follows)").Trim()
            ElseIf DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a clipboard object follows)").Trim()
                DoFileObjectClip = True
            Else
                DoFileObject = False
            End If

            ' Regular expression to find text in the format "(markup:..." and extract until ")"
            Dim pattern As String = Regex.Escape(TPMarkupTriggerL) & "(.*?)" & Regex.Escape(TPMarkupTriggerR)
            'Dim pattern As String = $"\{TPMarkupTriggerL}(.*?)\{TPMarkupTriggerR}"
            ' Match the pattern in the input string
            Dim match As Match = Regex.Match(OtherPrompt, pattern, RegexOptions.IgnoreCase)
            If match.Success Then
                ' Extract the captured group (the text between "(markup:" and ")")
                TPMarkupName = match.Groups(1).Value
                DoTPMarkup = True
                OtherPrompt = Regex.Replace(OtherPrompt, pattern, String.Empty, RegexOptions.IgnoreCase)
            End If

            If OtherPrompt.StartsWith(ClipboardPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix.Length).Trim()
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(ClipboardPrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix2.Length).Trim()
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(NewdocPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(NewdocPrefix.Length).Trim()
                DoClipboard = True
                DoChunks = False
                DoNewDoc = True
            ElseIf OtherPrompt.StartsWith(BubblesPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(BubblesPrefix.Length).Trim()
                DoBubbles = True
            ElseIf OtherPrompt.StartsWith(SlidesPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(SlidesPrefix.Length).Trim()
                DoSlides = True
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(InPlacePrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(InPlacePrefix.Length).Trim()
                DoInplace = True
            ElseIf OtherPrompt.StartsWith(AddPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(AddPrefix.Length).Trim()
                DoInplace = False
            ElseIf OtherPrompt.StartsWith(AddPrefix2, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(AddPrefix2.Length).Trim()
                DoInplace = False
            ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                DoMarkup = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefixRegex, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixRegex.Length).Trim()
                DoMarkup = True
                MarkupMethod = 4
            ElseIf OtherPrompt.StartsWith(MarkupPrefixWord, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixWord.Length).Trim()
                DoMarkup = True
                MarkupMethod = 1
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiffW, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiffW.Length).Trim()
                DoMarkup = True
                MarkupMethod = 3
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiff, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiff.Length).Trim()
                DoMarkup = True
                MarkupMethod = 2
            ElseIf OtherPrompt.StartsWith(PanePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PanePrefix.Length).Trim()
                DoPane = True
                DoClipboard = True
                DoChunks = False
            ElseIf OtherPrompt.StartsWith(PushbackPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PushbackPrefix.Length).Trim()
                DoPushback = True
                DoChunks = False
                DoBubblesExtract = True
            ElseIf OtherPrompt.StartsWith(PushbackPrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PushbackPrefix2.Length).Trim()
                DoPushback = True
                DoChunks = False
                DoBubblesExtract = True
            End If

            If OtherPrompt.IndexOf(NetTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NetTrigger, "").Trim()
                DoNet = True
            End If

            SelectedAlternateModels = Nothing
            If UseSecondAPI AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) AndAlso OtherPrompt.IndexOf(MultiModelTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                If Not DoMarkup AndAlso Not DoBubbles AndAlso Not DoPushback AndAlso Not DoSlides Then
                    If Not ShowMultipleModelSelection(_context, INI_AlternateModelPath) OrElse SelectedAlternateModels Is Nothing OrElse SelectedAlternateModels.Count = 0 Then
                        Return
                    End If
                Else
                    ShowCustomMessageBox($"The multi-model feature cannot be used together with markup, bubbles or slides - will continue only with the model you already selected.")
                End If
                OtherPrompt = OtherPrompt.Replace(MultiModelTrigger, "").Trim()
            End If

            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) And OtherPrompt.IndexOf(MyStyleTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)
                If Not IO.File.Exists(StylePath) Then
                    ShowCustomMessageBox("No MyStyle prompt file has been found. You may have to first create a MyStyle prompt. Go to 'Analyze' and use 'Define MyStyle' to do so - will abort.")
                    Return
                End If
                OtherPrompt = OtherPrompt.Replace(MyStyleTrigger, "").Trim()
                MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Word", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", False)
                If MyStyleInsert = "ERROR" Then Return
                If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then Return
                DoMyStyle = True
            End If

            If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(AddDocTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then

                InsertDocs = GatherSelectedDocuments()
                Debug.WriteLine($"GatherSelectedDocs returned: {Left(InsertDocs, 3000)}")
                If String.IsNullOrWhiteSpace(InsertDocs) Then
                    ShowCustomMessageBox("No content was found or an error occurred in gathering the additional document(s) - will abort.")
                    Return
                ElseIf InsertDocs.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then
                    ShowCustomMessageBox($"An error occured gathering the additional document(s) ({InsertDocs.Substring(6).Trim()}) - will abort.")
                    Return
                ElseIf InsertDocs.StartsWith("NONE", StringComparison.OrdinalIgnoreCase) Then
                    ShowCustomMessageBox($"There are no other documents to add - will abort.")
                    Return
                End If
                OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(AddDocTrigger), "", RegexOptions.IgnoreCase)
            End If

            If Not String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                ' Count total (case-insensitive) occurrences first
                Dim totalOccurrences As Integer =
                Regex.Matches(OtherPrompt, Regex.Escape(ExtTrigger), RegexOptions.IgnoreCase).Count

                ' Pattern that detects any tag that directly surrounds the placeholder, e.g. <tag>{doc}</tag>
                Dim wrappedPattern As String =
                "<(?<name>[A-Za-z][\w\-]*)\b[^>]*>\s*" & Regex.Escape(ExtTrigger) & "\s*</\k<name>>"

                If totalOccurrences = 1 Then
                    ' Original single-occurrence behavior (now with optional auto-wrapping)
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    doc = Await GetFileContent(Nothing, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object))
                    If String.IsNullOrWhiteSpace(doc) Then
                        ShowCustomMessageBox("The file you have selected is empty or not supported - will abort.")
                        Return
                    End If

                    Dim isWrapped As Boolean = Regex.IsMatch(OtherPrompt, wrappedPattern, RegexOptions.IgnoreCase)
                    Dim replacementText As String = If(isWrapped, doc, $"<document>{doc}</document>")

                    OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ExtTrigger), replacementText, RegexOptions.IgnoreCase)
                    ShowCustomMessageBox($"This file will be included in your prompt where you have referred to {ExtTrigger}: " & vbCrLf & vbCrLf & doc)

                Else
                    ' Multi-occurrence behavior: prompt separately per placeholder
                    For occurrence As Integer = 1 To totalOccurrences
                        Dim idx As Integer = OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase)
                        If idx < 0 Then Exit For

                        DragDropFormLabel = ""
                        DragDropFormFilter = ""
                        doc = Await GetFileContent(Nothing, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object))
                        If String.IsNullOrWhiteSpace(doc) Then
                            Dim answer As Integer = ShowCustomYesNoBox($"The file you selected for occurrence #{occurrence} is empty, not supported or you cancelled the upload. Do you want to continue or abort?", "Continue", "Abort")
                            If answer = 2 Then Return
                        End If

                        Dim replacementText As String = ""

                        If Not String.IsNullOrEmpty(doc) Then
                            ' Determine if this particular occurrence is already wrapped by any tag
                            Dim isWrappedThis As Boolean = False
                            Dim mcol As MatchCollection = Regex.Matches(OtherPrompt, wrappedPattern, RegexOptions.IgnoreCase)
                            For Each m As Match In mcol
                                If idx >= m.Index AndAlso idx < m.Index + m.Length Then
                                    isWrappedThis = True
                                    Exit For
                                End If
                            Next

                            replacementText = If(isWrappedThis, doc, $"<document{occurrence}>{doc}</document{occurrence}>")

                        End If

                        ' Replace only the first remaining occurrence manually
                        OtherPrompt = OtherPrompt.Substring(0, idx) &
                                  replacementText &
                                  OtherPrompt.Substring(idx + ExtTrigger.Length)

                        If Not String.IsNullOrWhiteSpace(doc) Then
                            ShowCustomMessageBox($"This file will be included at occurrence #{occurrence} (of {totalOccurrences}) where you used {ExtTrigger}:" &
                                         vbCrLf & vbCrLf & doc)
                        End If
                    Next
                End If
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
                        Return
                    End If
                End If
            End If

            If DoSlides Then
                DragDropFormLabel = "A Powerpoint (pptx) file (or cancel to create one)."
                DragDropFormFilter = "Supported Files|*.pptx"
                SlideDeck = GetFileName()
                DragDropFormLabel = ""
                DragDropFormFilter = ""
                If String.IsNullOrWhiteSpace(SlideDeck) Then

                    ' Ask user first
                    Dim CreatePPTX As Integer = ShowCustomYesNoBox(
                         "You have not provided a Powerpoint file. Do you want create a new one?", "Yes", "No, abort")
                    If CreatePPTX <> 1 Then
                        ShowCustomMessageBox("No Powerpoint file has been selected - will abort. You can try again (use Ctrl-P to re-insert your prompt).")
                        Return
                    End If

                    ' If SlideDeck is empty, default to Desktop\NewPresentation.pptx
                    If String.IsNullOrWhiteSpace(SlideDeck) Then
                        Dim desktop As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
                        SlideDeck = System.IO.Path.Combine(desktop, "NewPresentation.pptx")
                    End If

                    ' Ensure we do not overwrite an existing file -> NewPresentation (2).pptx, (3).pptx, ...
                    If System.IO.File.Exists(SlideDeck) Then
                        Dim dir As String = System.IO.Path.GetDirectoryName(SlideDeck)
                        Dim name As String = System.IO.Path.GetFileNameWithoutExtension(SlideDeck)
                        Dim ext As String = System.IO.Path.GetExtension(SlideDeck)
                        Dim i As Integer = 2
                        Do
                            Dim candidate As String = System.IO.Path.Combine(dir, name & " (" & i.ToString() & ")" & ext)
                            If Not System.IO.File.Exists(candidate) Then
                                SlideDeck = candidate
                                Exit Do
                            End If
                            i += 1
                        Loop
                    End If

                    ' Create + save while PowerPoint is visible; then close everything cleanly
                    Dim pptApp As NetOffice.PowerPointApi.Application = Nothing
                    Dim presentation As NetOffice.PowerPointApi.Presentation = Nothing

                    Try
                        pptApp = New NetOffice.PowerPointApi.Application()

                        ' Visible on purpose (requested). Suppress alerts to avoid prompts.
                        pptApp.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                        pptApp.DisplayAlerts = NetOffice.PowerPointApi.Enums.PpAlertLevel.ppAlertsNone
                        pptApp.WindowState = NetOffice.PowerPointApi.Enums.PpWindowState.ppWindowNormal

                        ' Create a new presentation with a window
                        presentation = pptApp.Presentations.Add(NetOffice.OfficeApi.Enums.MsoTriState.msoTrue)

                        ' Ensure at least one slide exists
                        'presentation.Slides.Add(1, NetOffice.PowerPointApi.Enums.PpSlideLayout.ppLayoutBlank)

                        ' Save explicitly as Open XML (.pptx)
                        presentation.SaveAs(
                            SlideDeck,
                            NetOffice.PowerPointApi.Enums.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                            NetOffice.OfficeApi.Enums.MsoTriState.msoFalse
                        )

                        ' Close presentation (no prompts) and quit PowerPoint
                        presentation.Close()
                        pptApp.Quit()

                    Catch comEx As System.Runtime.InteropServices.COMException
                        ShowCustomMessageBox("PowerPoint COM error while creating file:" & vbCrLf &
                     "Message: " & comEx.Message & vbCrLf &
                     "HResult: 0x" & comEx.HResult.ToString("X8"))
                        Return

                    Catch ex As System.Exception
                        ShowCustomMessageBox("Error while creating PowerPoint file: " & ex.ToString())
                        Return

                    Finally
                        ' Dispose COM objects in correct order, even on error
                        If presentation IsNot Nothing Then
                            Try : presentation.Dispose() : Catch : End Try
                            presentation = Nothing
                        End If
                        If pptApp IsNot Nothing Then
                            Try : pptApp.Quit() : Catch : End Try   ' in case we failed before quitting
                            Try : pptApp.Dispose() : Catch : End Try
                            pptApp = Nothing
                        End If
                    End Try


                End If
            End If


            If NoText AndAlso (DoBubbles Or DoChunks) Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Ask the LLM to comment on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Return
                End If
            End If

            If NoText AndAlso DoMarkup Then
                Dim FullDocument As Integer = ShowCustomYesNoBox("You have not selected text. Do the markup on the full document?", "Yes", "No, abort")
                If FullDocument = 1 Then
                    Dim document As Word.Document = application.ActiveDocument
                    document.Content.Select()
                    NoText = False
                Else
                    Return
                End If
            End If

            If Not DoInplace AndAlso DoMarkup Then
                Dim AppendMarkup As Integer = ShowCustomYesNoBox("You have asked for a markup to be created, but according to the configuration, it will not replace your current selection but added to it at the end. Is this really what you want?", "Yes, add markup ", "No, replace text with markup")
                If AppendMarkup = 0 Then
                    Return
                ElseIf AppendMarkup = 2 Then
                    DoInplace = True
                    NoFormatAndFieldSaving = False
                End If
            End If

            If OtherPrompt.StartsWith(PurePrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(PurePrefix.Length).Replace("(a file object follows)", "").Replace("(a clipboard object follows)", "").Trim()
                SysPrompt = OtherPrompt
            Else
                If DoLib Then
                    Dim isSuccess As Boolean = Await ConsultLibrary(DoMarkup) ' updates SysPrompt
                    If Not isSuccess Then Return
                ElseIf DoNet Then
                    Dim isSuccess As Boolean = Await ConsultInternet(DoMarkup) ' updates SysPrompt
                    If Not isSuccess Then Return
                ElseIf NoText Then
                    SysPrompt = SP_FreestyleNoText
                Else
                    SysPrompt = SP_FreestyleText
                    If DoBubbles Then SysPrompt = SysPrompt & " " & SP_Add_Bubbles
                    If DoPushback Then SysPrompt = SysPrompt & " " & SP_Add_BubblesReply
                    If INI_MarkdownBubbles Then FormatInstruction = SP_Add_Bubbles_Format Else FormatInstruction = ""
                End If
            End If

            If DoChunks Then
                Dim response As String = SLib.ShowCustomInputBox($"How many paragraphs shall be treated at the same time (max. 25)?", "Iterate through the text", True, ChunkSize.ToString()).Trim()
                If Not Integer.TryParse(response, ChunkSize) Then ChunkSize = 0
                If response = "" OrElse response.ToLower() = "esc" OrElse ChunkSize = 0 Then Return
                If ChunkSize > 25 Then ChunkSize = 25
            Else
                ChunkSize = 0
            End If

            Debug.WriteLine("Freestyle Prompt: " & SysPrompt)

            Dim result As String = Await ProcessSelectedText(InterpolateAtRuntime(SysPrompt), True, DoKeepFormat, DoKeepParaFormat, DoInplace, DoMarkup, MarkupMethod, DoClipboard, DoBubbles, False, UseSecondAPI, KeepFormatCap, DoTPMarkup, TPMarkupName, False, FileObject, DoPane, ChunkSize, NoFormatAndFieldSaving, DoNewDoc, SlideDeck, InsertDocs <> "", DoMyStyle, DoBubblesExtract, DoPushback)

            If UseSecondAPI And originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



End Class
