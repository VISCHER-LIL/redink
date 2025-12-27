
' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Explicit On
Option Strict On

Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Public NotInheritable Class IniImportManager

    Private Sub New()
    End Sub

    ' -------------------------------
    ' CONSTANTS (as requested)
    ' -------------------------------
    Private Const MAX_DOWNLOAD_BYTES As System.Int32 = 50 * 1024

    ' Put your own domains here. Comparison is ordinal-ignore-case.
    Private Shared ReadOnly ALLOWED_HOSTS As System.String() = New System.String() {
        "redink.ai",
        "www.rosenthal.ch",
        "lexisearch.ch"
    }

    Private Const TITLE_IMPORT As System.String = AN & " Settings Import"

    Private Enum ImportKind
        PrimaryModel = 1
        SecondaryModel = 2
        AlternateModel = 3
        SpecialService = 4
        OtherParameters = 5
    End Enum

    ' -------------------------------
    ' PUBLIC ENTRY POINT
    ' -------------------------------
    Public Shared Function RunImportFromVariableConfigurationWindow(context As ISharedContext, ownerForm As System.Windows.Forms.Form) As Boolean

        Dim mainIniChanged As System.Boolean = False


        If context Is Nothing Then
            ShowCustomMessageBox("Internal error: context is missing.")
            Return False
        End If

        Dim activeIniPath As System.String = Nothing
        Try
            activeIniPath = GetActiveConfigFilePath(context)
        Catch ex As System.Exception
            ShowCustomMessageBox("Could not determine active configuration file path: " & ex.Message)
            Return False
        End Try

        If System.String.IsNullOrWhiteSpace(activeIniPath) Then
            ShowCustomMessageBox("No active configuration file path found.")
            Return False
        End If

        ' Enforce: only when redink.ini is local per-application (not registry-priority and not Excel/Outlook using Word)
        Dim disableReason As System.String = Nothing
        If Not CanUseImportFeature(context, activeIniPath, disableReason) Then
            ShowCustomMessageBox(disableReason)
            Return False
        End If

        If Not System.IO.File.Exists(activeIniPath) Then
            ShowCustomMessageBox("The main configuration file does not exist: " & activeIniPath)
            Return False
        End If

        ' 1) Ask for URL or file path
        Dim sourceText As System.String = Nothing
        Dim sourceLabel As System.String = Nothing

        If Not TryGetImportSourceText(ownerForm, sourceText, sourceLabel) Then
            Return False ' user aborted or failed
        End If

        If System.String.IsNullOrWhiteSpace(sourceText) Then
            ShowCustomMessageBox("No content loaded.")
            Return False
        End If

        ' 2) Show content to user (viewer)

        Dim editedText As System.String = Nothing

        If Not ShowTextAsViewer(ownerForm, sourceText, "Import Source", editedText) OrElse String.IsNullOrEmpty(editedText) Then
            ShowCustomMessageBox("Import aborted.")
            Return False ' user cancelled
        End If

        sourceText = editedText


        ' 3) Decide import kind
        Dim kind As ImportKind

        Dim hostForm As System.Windows.Forms.Form = TryCast(ownerForm, System.Windows.Forms.Form)
        Dim wasTopMost As System.Boolean = False
        Dim hadHost As System.Boolean = (hostForm IsNot Nothing)

        If hadHost Then
            wasTopMost = hostForm.TopMost
            hostForm.TopMost = False
            hostForm.Enabled = False
            System.Windows.Forms.Application.DoEvents()
        End If

        Try
            If Not TryChooseImportKind(ownerForm, kind) Then
                Return False
            End If
        Finally
            If hadHost Then
                hostForm.Enabled = True
                hostForm.TopMost = wasTopMost
                hostForm.Activate()
            End If
        End Try


        ' 4) Normalize import text:
        '    - If kind is Primary/Secondary/OtherParameters: drop section headers and only keep non-section lines
        Dim normalizedImportText As System.String = NormalizeImportTextForKind(sourceText, kind)

        ' 5) Placeholder capture and substitution
        Dim substitutedText As System.String = normalizedImportText
        Dim placeholderWarnings As New System.Collections.Generic.List(Of System.String)()

        If Not TryResolvePlaceholders(ownerForm, substitutedText, placeholderWarnings) Then
            Return False ' user aborted the placeholder dialog
        End If

        If placeholderWarnings.Count > 0 Then
            ShowCustomMessageBox(System.String.Join(System.Environment.NewLine & System.Environment.NewLine, placeholderWarnings))
        End If

        ' 6) Determine target ini + optional section name
        Dim mainIniPath As System.String = activeIniPath
        Dim altIniPath As System.String = Nothing
        Dim svcIniPath As System.String = Nothing
        Dim targetIniPath As System.String = Nothing
        Dim targetSectionName As System.String = Nothing

        If kind = ImportKind.AlternateModel Then
            Dim pathUpdated As System.Boolean = False

            If Not TryEnsureSectionedIniPath(ownerForm,
                                 context,
                                 mainIniPath,
                                 isAlternate:=True,
                                 targetIniPath:=altIniPath,
                                 mainIniWasUpdated:=pathUpdated) Then
                Return False
            End If

            If pathUpdated Then
                mainIniChanged = True
            End If

            targetIniPath = altIniPath
            If Not TryGetSectionNameFromImportText(ownerForm, substitutedText, targetSectionName, "alternate model") Then
                Return False
            End If

        ElseIf kind = ImportKind.SpecialService Then
            Dim pathUpdated As System.Boolean = False

            If Not TryEnsureSectionedIniPath(ownerForm,
                                 context,
                                 mainIniPath,
                                 isAlternate:=False,
                                 targetIniPath:=altIniPath,
                                 mainIniWasUpdated:=pathUpdated) Then
                Return False
            End If

            If pathUpdated Then
                mainIniChanged = True
            End If
            targetIniPath = svcIniPath
            If Not TryGetSectionNameFromImportText(ownerForm, substitutedText, targetSectionName, "special service") Then
                Return False
            End If

        Else
            targetIniPath = mainIniPath
        End If

        ' 7) Parse imported key/value lines (for redink updates) or section body (for sectioned)
        Dim importedLines As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserveNonEmpty(substitutedText)

        If importedLines.Count = 0 Then
            ShowCustomMessageBox("The import content is empty after processing.")
            Return False
        End If

        ' 8) Build dry-run plan(s)
        Dim plans As New System.Collections.Generic.List(Of DryRunPlan)
        Dim allRemovedLines As New System.Collections.Generic.List(Of System.String)

        Try
            If kind = ImportKind.AlternateModel OrElse kind = ImportKind.SpecialService Then

                Dim segments = ParseIniSegments(importedLines)

                If segments.Count = 0 Then
                    Throw New System.Exception("No valid sections found.")
                End If

                For Each kvp In segments
                    Dim sectionName As System.String = kvp.Key
                    Dim sectionLines As System.Collections.Generic.List(Of System.String) = kvp.Value

                    Dim plan As DryRunPlan =
                BuildDryRunPlan(context, kind, targetIniPath, sectionName, sectionLines)

                    If plan.WillCreateRemovedBackup Then
                        allRemovedLines.AddRange(plan.RemovedLinesBackup)
                        plan.WillCreateRemovedBackup = False
                    End If

                    plans.Add(plan)
                Next

            Else
                ' Non-sectioned import: unchanged behavior
                Dim singlePlan As DryRunPlan =
            BuildDryRunPlan(context, kind, targetIniPath, targetSectionName, importedLines)

                If singlePlan.WillCreateRemovedBackup Then
                    allRemovedLines.AddRange(singlePlan.RemovedLinesBackup)
                    singlePlan.WillCreateRemovedBackup = False
                End If

                plans.Add(singlePlan)
            End If

        Catch ex As System.Exception
            ShowCustomMessageBox("Could not build dry run plan: " & ex.Message)
            Return False
        End Try

        If plans.Count = 0 Then
            ShowCustomMessageBox("Nothing to import.")
            Return False
        End If

        ' 9) Composite dry-run summary
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("Dry run – review before importing")
        sb.AppendLine()

        For Each p As DryRunPlan In plans
            sb.AppendLine("Target file: " & p.TargetIniPath)
            If Not System.String.IsNullOrWhiteSpace(p.TargetSectionName) Then
                sb.AppendLine("Section: [" & p.TargetSectionName & "]")
            End If
            If p.OverwrittenKeys.Count > 0 Then
                sb.AppendLine("Overwritten keys:")
                For Each k As System.String In p.OverwrittenKeys
                    sb.AppendLine("  - " & k)
                Next
            End If
            sb.AppendLine()
        Next

        sb.AppendLine("A full backup will be created.")
        If allRemovedLines.Count > 0 Then
            sb.AppendLine("All removed content will be stored in ONE backup file.")
        End If
        sb.AppendLine()
        sb.AppendLine("Proceed with import?")

        Dim decision As System.Int32 = ShowCustomYesNoBox(sb.ToString(), "Yes, continue", "No, abort import")

        If decision <> 1 Then Return False

        ' 10) Commit all plans (in order)
        For Each p As DryRunPlan In plans
            CommitDryRunPlan(p)
        Next

        For Each p As DryRunPlan In plans
            If System.String.Equals(p.TargetIniPath,
                            mainIniPath,
                            System.StringComparison.OrdinalIgnoreCase) Then
                mainIniChanged = True
                Exit For
            End If
        Next

        ' Write ONE combined removed-content backup
        If allRemovedLines.Count > 0 Then
            Dim ts As System.String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim baseName As System.String = System.IO.Path.GetFileNameWithoutExtension(targetIniPath)
            Dim dir As System.String = System.IO.Path.GetDirectoryName(targetIniPath)

            Dim removedBackup As System.String =
        System.IO.Path.Combine(dir, baseName & "_removed_" & ts & ".bak")

            System.IO.File.WriteAllText(
        removedBackup,
        System.String.Join(System.Environment.NewLine, allRemovedLines),
        System.Text.Encoding.UTF8
    )
        End If

        If Not mainIniChanged Then ShowCustomMessageBox("Import completed. No reloading or restarting required.")

        Return mainIniChanged

    End Function

    ' =========================================================================================
    '  INTEGRATION INTO ShowVariableConfigurationWindow(): Add a new button like btnEditIni
    ' =========================================================================================
    Public Shared Sub IntegrateButtonIntoVariableConfigurationWindow(form As System.Windows.Forms.Form,
                                                                     pnlButtons As System.Windows.Forms.FlowLayoutPanel,
                                                                     context As ISharedContext)

        If form Is Nothing OrElse pnlButtons Is Nothing Then Return

        Dim btnImport As New System.Windows.Forms.Button() With {
            .Text = "Import Settings",
            .AutoSize = True,
            .Margin = New System.Windows.Forms.Padding(10)
        }

        ' Enable/disable according to your rules
        Try
            Dim activeIniPath As System.String = GetActiveConfigFilePath(context)
            Dim reason As System.String = Nothing
            Dim canUse As System.Boolean = CanUseImportFeature(context, activeIniPath, reason)
            btnImport.Enabled = canUse
            If Not canUse AndAlso Not System.String.IsNullOrWhiteSpace(reason) Then
                btnImport.Tag = reason
                AddHandler btnImport.MouseHover,
                    Sub()
                        ' If you have a helper tooltip, use it; otherwise ignore
                    End Sub
            End If
        Catch
            btnImport.Enabled = False
        End Try

        AddHandler btnImport.Click,
            Sub()
                ' Z-order fix similar to your editor integration
                Dim wasTopMost As System.Boolean = form.TopMost
                form.TopMost = False
                form.Enabled = False
                System.Windows.Forms.Application.DoEvents()

                Try
                    IniImportManager.RunImportFromVariableConfigurationWindow(context, form)
                Finally
                    form.Enabled = True
                    form.TopMost = wasTopMost
                    form.Activate()
                End Try
            End Sub

        ' Add next to the edit button (or wherever you prefer)
        pnlButtons.Controls.Add(btnImport)

    End Sub

    ' =========================================================================================
    '  FEATURE ENABLEMENT RULES
    ' =========================================================================================
    Private Shared Function CanUseImportFeature(context As ISharedContext,
                                               activeIniPath As System.String,
                                               ByRef disableReason As System.String) As System.Boolean

        disableReason = Nothing

        If context Is Nothing Then
            disableReason = "Import is not available (missing context)."
            Return False
        End If

        If System.String.IsNullOrWhiteSpace(activeIniPath) Then
            disableReason = "Import is not available (no active .ini path)."
            Return False
        End If

        ' Rule 1: If redink.ini loaded from registry path with priority => disable.
        ' We cannot reliably infer registry-priority solely from a file path without your internal flags.
        ' Best effort: if you have a context flag indicating registry-priority path in effect, check it here.
        ' If you do not, you should expose one; otherwise you risk enabling import incorrectly.
        Try
            If RegPath_IniPrio Then
                disableReason = "Import is not available when the configuration is controlled via registry/network setup."
                Return False
            End If
        Catch
            ' If flag not available, ignore; but you should add it.
        End Try

        ' Rule 2: If Excel/Outlook uses Word's path => disable in Excel/Outlook.
        Dim rdv As System.String = Nothing
        Try
            rdv = context.RDV
        Catch
        End Try

        If Not System.String.IsNullOrWhiteSpace(rdv) Then
            Dim defaultPathThisApp As System.String = Nothing
            Dim defaultPathWord As System.String = Nothing

            Try
                defaultPathThisApp = GetDefaultINIPath(rdv)
            Catch
            End Try

            Try
                defaultPathWord = GetDefaultINIPath("Word")
            Catch
            End Try

            If (System.String.Equals(rdv, "Excel", System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(rdv, "Outlook", System.StringComparison.OrdinalIgnoreCase)) AndAlso
               Not System.String.IsNullOrWhiteSpace(defaultPathWord) AndAlso
               System.IO.File.Exists(defaultPathWord) AndAlso
               System.String.Equals(activeIniPath, defaultPathWord, System.StringComparison.OrdinalIgnoreCase) Then

                disableReason = "Import is not available here because this application is using Word's configuration file. Please use Word to import settings."
                Return False
            End If

            ' Only enable if active path equals this app's own default path (local per-application)
            If Not System.String.IsNullOrWhiteSpace(defaultPathThisApp) AndAlso
               System.IO.File.Exists(defaultPathThisApp) Then

                If Not System.String.Equals(activeIniPath, defaultPathThisApp, System.StringComparison.OrdinalIgnoreCase) Then
                    disableReason = "Import is only available when using the local per-application configuration file."
                    Return False
                End If
            End If
        End If

        Return True

    End Function

    ' =========================================================================================
    '  SOURCE ACQUISITION (URL or FILE)
    ' =========================================================================================
    Private Shared Function TryGetImportSourceText(ownerForm As System.Windows.Forms.Form,
                                                  ByRef sourceText As System.String,
                                                  ByRef sourceLabel As System.String) As System.Boolean

        sourceText = Nothing
        sourceLabel = Nothing

        Dim input As String = ShowCustomInputBox("Enter the source URL (https://...) or file / UNC path:", TITLE_IMPORT, True)

        If System.String.IsNullOrWhiteSpace(input) Then
            ShowCustomMessageBox("No source provided.")
            Return False
        End If

        input = input.Trim()

        If input.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then

            Dim u As System.Uri = Nothing
            Try
                u = New System.Uri(input)
            Catch
                ShowCustomMessageBox("Invalid URL.")
                Return False
            End Try

            ' Allowlist check (warn on non-allowlisted)
            Dim isAllowed As System.Boolean = IsHostAllowed(u.Host)
            If Not isAllowed Then
                Dim warnDecision As System.Int32 = ShowCustomYesNoBox(
                    "Warning: This URL host is not on the built-in trust list:" & System.Environment.NewLine &
                    u.Host & System.Environment.NewLine & System.Environment.NewLine &
                    "Importing configuration from unknown hosts can be dangerous." & System.Environment.NewLine & System.Environment.NewLine &
                    "Do you want to continue?",
                    "Yes, continue",
                    "No, abort import"
                )
                If warnDecision <> 1 Then
                    Return False
                End If
            End If

            Dim downloaded As System.String = Nothing
            Try
                downloaded = DownloadHttpsTextWithLimit(u, MAX_DOWNLOAD_BYTES)
            Catch ex As System.Exception
                ShowCustomMessageBox("Download failed: " & ex.Message)
                Return False
            End Try

            sourceText = downloaded
            sourceLabel = u.ToString()
            Return True

        Else
            ' Local / UNC path
            Dim path As System.String = input

            ' If user entered env vars, expand them
            Try
                path = ExpandEnvironmentVariables(path)
            Catch
            End Try

            If Not System.IO.File.Exists(path) Then
                ShowCustomMessageBox("File not found: " & path)
                Return False
            End If

            Try
                Dim fi As New System.IO.FileInfo(path)
                If fi.Length > MAX_DOWNLOAD_BYTES Then
                    ShowCustomMessageBox("The file is larger than the allowed limit (" & MAX_DOWNLOAD_BYTES.ToString() & " bytes).")
                    Return False
                End If
            Catch
                ' ignore
            End Try

            Try
                sourceText = System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8)
            Catch ex As System.Exception
                ShowCustomMessageBox("Could not read file: " & ex.Message)
                Return False
            End Try

            sourceLabel = path
            Return True
        End If

    End Function

    Private Shared Function IsHostAllowed(host As System.String) As System.Boolean
        If System.String.IsNullOrWhiteSpace(host) Then Return False
        For Each h As System.String In ALLOWED_HOSTS
            If System.String.Equals(host, h, System.StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    Private Shared Function DownloadHttpsTextWithLimit(u As System.Uri, maxBytes As System.Int32) As System.String

        If u Is Nothing Then Throw New System.ArgumentNullException(NameOf(u))
        If Not System.String.Equals(u.Scheme, "https", System.StringComparison.OrdinalIgnoreCase) Then
            Throw New System.Exception("Only HTTPS is allowed.")
        End If

        ' Explicitly enable TLS
        Try
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls12 Or
                CType(0, System.Net.SecurityProtocolType) ' keep compiler happy; no-op
        Catch
            ' ignore
        End Try

        Dim req As System.Net.HttpWebRequest = CType(System.Net.WebRequest.Create(u), System.Net.HttpWebRequest)
        req.Method = "GET"
        req.AllowAutoRedirect = True
        req.UserAgent = "RedInk-Importer"
        req.Timeout = 15000
        req.ReadWriteTimeout = 15000

        Using resp As System.Net.HttpWebResponse = CType(req.GetResponse(), System.Net.HttpWebResponse)
            Using s As System.IO.Stream = resp.GetResponseStream()
                If s Is Nothing Then Throw New System.Exception("No response stream.")
                Using ms As New System.IO.MemoryStream()
                    Dim buffer(4095) As System.Byte
                    Dim total As System.Int32 = 0
                    While True
                        Dim read As System.Int32 = s.Read(buffer, 0, buffer.Length)
                        If read <= 0 Then Exit While
                        total += read
                        If total > maxBytes Then
                            Throw New System.Exception("Download exceeds the maximum allowed size (" & maxBytes.ToString() & " bytes).")
                        End If
                        ms.Write(buffer, 0, read)
                    End While
                    Dim data As System.Byte() = ms.ToArray()
                    Return System.Text.Encoding.UTF8.GetString(data)
                End Using
            End Using
        End Using

    End Function

    Private Shared Function ShowTextAsViewer(
                            ownerForm As System.Windows.Forms.Form,
                            text As System.String,
                            title As System.String,
                            ByRef finalText As System.String
                        ) As Boolean

        Dim tmp As System.String =
            System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "RedInk_ImportPreview_" & System.Guid.NewGuid().ToString("N") & ".txt"
            )

        System.IO.File.WriteAllText(tmp, text, System.Text.Encoding.UTF8)

        Try
            Dim wasSaved As System.Boolean? = Nothing
            ShowTextFileEditor(tmp, title, False, Nothing, wasSaved)

            If wasSaved.HasValue AndAlso wasSaved.Value Then
                ' Re-read edited file
                finalText = System.IO.File.ReadAllText(tmp, System.Text.Encoding.UTF8)
                Return True
            End If

            ' Cancel / close
            finalText = Nothing
            Return False

        Finally
            Try
                System.IO.File.Delete(tmp)
            Catch
            End Try
        End Try

    End Function

    ' =========================================================================================
    '  IMPORT KIND SELECTION
    ' =========================================================================================
    Private Shared Function TryChooseImportKind(ownerForm As System.Windows.Forms.Form, ByRef kind As ImportKind) As System.Boolean
        kind = ImportKind.PrimaryModel

        Dim options As New System.Collections.Generic.List(Of System.String)() From {
            "For the primary model",
            "For the secondary model",
            "For an alternate model",
            "For a special service",
            "For other parameters"
        }

        Dim choice As System.String = ShowSelectionForm("Which settings do you want to import?", TITLE_IMPORT, options)
        If System.String.IsNullOrWhiteSpace(choice) Then Return False

        If choice.StartsWith("For the primary", System.StringComparison.OrdinalIgnoreCase) Then
            kind = ImportKind.PrimaryModel
        ElseIf choice.StartsWith("For the secondary", System.StringComparison.OrdinalIgnoreCase) Then
            kind = ImportKind.SecondaryModel
        ElseIf choice.StartsWith("For an alternate", System.StringComparison.OrdinalIgnoreCase) Then
            kind = ImportKind.AlternateModel
        ElseIf choice.StartsWith("For a special service", System.StringComparison.OrdinalIgnoreCase) Then
            kind = ImportKind.SpecialService
        Else
            kind = ImportKind.OtherParameters
        End If

        Return True
    End Function

    Private Shared Function NormalizeImportTextForKind(sourceText As System.String, kind As ImportKind) As System.String
        If System.String.IsNullOrWhiteSpace(sourceText) Then Return ""

        If kind = ImportKind.PrimaryModel OrElse kind = ImportKind.SecondaryModel OrElse kind = ImportKind.OtherParameters Then
            ' Drop any [section] headers entirely, keep all other lines unchanged.
            Dim lines As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserve(sourceText)
            Dim kept As New System.Collections.Generic.List(Of System.String)()

            For Each line As System.String In lines
                Dim t As System.String = line.Trim()
                If t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso t.Length >= 2 Then
                    ' drop header
                    Continue For
                End If
                kept.Add(line)
            Next

            Return System.String.Join(System.Environment.NewLine, kept)
        End If

        Return sourceText
    End Function

    ' =========================================================================================
    '  PLACEHOLDERS [[...]] -> prompt user
    ' =========================================================================================
    Private Shared Function TryResolvePlaceholders(ownerForm As System.Windows.Forms.Form,
                                                  ByRef text As System.String,
                                                  warnings As System.Collections.Generic.List(Of System.String)) As System.Boolean

        If warnings Is Nothing Then warnings = New System.Collections.Generic.List(Of System.String)()
        If System.String.IsNullOrWhiteSpace(text) Then Return True

        Dim rx As New System.Text.RegularExpressions.Regex("\[\[(.+?)\]\]",
                                                          System.Text.RegularExpressions.RegexOptions.Singleline)

        Dim matches As System.Text.RegularExpressions.MatchCollection = rx.Matches(text)
        If matches Is Nothing OrElse matches.Count = 0 Then Return True

        Dim unique As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        For Each m As System.Text.RegularExpressions.Match In matches
            If m Is Nothing OrElse Not m.Success Then Continue For
            Dim key As System.String = m.Groups(1).Value
            If System.String.IsNullOrWhiteSpace(key) Then Continue For
            If Not unique.ContainsKey(key) Then unique.Add(key, "")
        Next

        If unique.Count = 0 Then Return True

        Dim paramList As New System.Collections.Generic.List(Of InputParameter)()
        For Each k As System.String In unique.Keys
            paramList.Add(New InputParameter(k, ""))
        Next

        Dim params() As InputParameter = paramList.ToArray()

        If ShowCustomVariableInputForm("The settings require your to enter individual values. Please enter them (leave empty to keep a placeholder and edit later): ",
                                       TITLE_IMPORT,
                                       params) = False Then
            Return False
        End If

        ' Replace if non-empty, else keep placeholder and warn
        For Each p As InputParameter In params
            Dim name As System.String = System.Convert.ToString(p.Name)
            Dim value As System.String = System.Convert.ToString(p.Value)

            If System.String.IsNullOrWhiteSpace(name) Then Continue For

            If Not System.String.IsNullOrWhiteSpace(value) Then
                ' Replace all occurrences, preserving original placeholder text format [[name]] case-insensitively
                ' We'll replace using regex with escaped key inside.
                Dim keyRx As New System.Text.RegularExpressions.Regex("\[\[" & System.Text.RegularExpressions.Regex.Escape(name) & "\]\]",
                                                                     System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                text = keyRx.Replace(text, value)
            Else
                warnings.Add("Warning: Placeholder '[[ " & name & " ]]' was left empty and remains in the configuration." & System.Environment.NewLine &
                             "You can later fill it using the 'Edit .ini Files' feature or directly access the file.")
            End If
        Next

        Return True
    End Function

    ' =========================================================================================
    '  ENSURE AlternateModelPath / SpecialServicePath exist in redink.ini (ask user, confirm, optional create file)
    ' =========================================================================================
    Private Shared Function TryEnsureSectionedIniPath(
                                        ownerForm As System.Windows.Forms.Form,
                                        context As ISharedContext,
                                        mainIniPath As System.String,
                                        isAlternate As System.Boolean,
                                        ByRef targetIniPath As System.String,
                                        Optional ByRef mainIniWasUpdated As System.Boolean = False
                                    ) As System.Boolean


        targetIniPath = Nothing

        Dim currentSetting As System.String = Nothing
        Dim settingKey As System.String = If(isAlternate, "AlternateModelPath", "SpecialServicePath")
        Dim defaultFileName As System.String = If(isAlternate, "allmodels.ini", "specialservices.ini")

        Try
            currentSetting = If(isAlternate, context.INI_AlternateModelPath, context.INI_SpecialServicePath)
        Catch
        End Try

        Dim expandedCurrent As System.String = Nothing
        If Not System.String.IsNullOrWhiteSpace(currentSetting) Then
            Try
                expandedCurrent = ExpandEnvironmentVariables(currentSetting)
            Catch
                expandedCurrent = currentSetting
            End Try
        End If

        If Not System.String.IsNullOrWhiteSpace(expandedCurrent) Then
            targetIniPath = expandedCurrent
        Else
            ' Ask user for path (default: same directory as main ini)
            Dim baseDir As System.String = System.IO.Path.GetDirectoryName(mainIniPath)
            Dim suggested As System.String = System.IO.Path.Combine(baseDir, defaultFileName)

            Dim p0 As InputParameter = New InputParameter(settingKey & " (file path)", suggested)

            ' Offer portable storage if user wants and if we can express it as %APPSTARTUPPATH%
            Dim canUseAppStartupToken As System.Boolean = False
            Dim appStartup As System.String = Nothing
            Try
                appStartup = System.Windows.Forms.Application.StartupPath
                If Not System.String.IsNullOrWhiteSpace(appStartup) AndAlso
                   System.String.Equals(System.IO.Path.GetFullPath(appStartup).TrimEnd("\"c),
                                       System.IO.Path.GetFullPath(baseDir).TrimEnd("\"c),
                                       System.StringComparison.OrdinalIgnoreCase) Then
                    canUseAppStartupToken = True
                End If
            Catch
            End Try

            Dim chosenPath As String =
                        ShowCustomInputBox(
                            "Please confirm or change the file path for " & settingKey & " (if unsure, just confirm):",
                            TITLE_IMPORT,
                            True,
                            suggested
                        )

            If System.String.IsNullOrWhiteSpace(chosenPath) Then
                ShowCustomMessageBox("No path provided. Import aborted.")
                Return False
            End If

            chosenPath = chosenPath.Trim()

            ' Expand (for access), but store per user preference
            Dim expandedChosen As System.String = chosenPath
            Try
                expandedChosen = ExpandEnvironmentVariables(chosenPath)
            Catch
            End Try

            targetIniPath = expandedChosen

            ' Confirm create file if missing (abort if user does not want creation)
            If Not System.IO.File.Exists(targetIniPath) Then
                Dim createDecision As System.Int32 = ShowCustomYesNoBox(
                    "The file does not exist:" & System.Environment.NewLine & targetIniPath & System.Environment.NewLine & System.Environment.NewLine &
                    "Do you want to create it now? If you choose No, the import will abort (if unsure, choose Yes).",
                    "Yes, create file",
                    "No, abort import"
                )
                If createDecision <> 1 Then
                    Return False
                End If

                Try
                    Dim dir As System.String = System.IO.Path.GetDirectoryName(targetIniPath)
                    If Not System.String.IsNullOrWhiteSpace(dir) AndAlso Not System.IO.Directory.Exists(dir) Then
                        System.IO.Directory.CreateDirectory(dir)
                    End If
                    System.IO.File.WriteAllText(targetIniPath,
                                                "; created by Red Ink Settings Importer on " & System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & System.Environment.NewLine,
                                                System.Text.Encoding.UTF8)
                Catch ex As System.Exception
                    ShowCustomMessageBox("Could not create file: " & ex.Message)
                    Return False
                End Try
            End If

            ' Store the path into main redink.ini (with optional %APPSTARTUPPATH% style)
            Dim valueToStore As System.String = chosenPath

            ' Write settingKey=valueToStore into main ini (line preserving, overwrite if exists, else append)
            Try
                Dim plan As DryRunPlan = BuildDryRunPlanForSingleKey(mainIniPath, settingKey, valueToStore)
                CommitDryRunPlan(plan)
                mainIniWasUpdated = True
            Catch ex As System.Exception
                ShowCustomMessageBox("Could not update main configuration with " & settingKey & ": " & ex.Message)
                Return False
            End Try
        End If

        Return True

    End Function

    ' =========================================================================================
    '  SECTION NAME EXTRACTION (Alternate/Special)
    ' =========================================================================================
    Private Shared Function TryGetSectionNameFromImportText(ownerForm As System.Windows.Forms.Form,
                                                           text As System.String,
                                                           ByRef sectionName As System.String,
                                                           friendlyType As System.String) As System.Boolean

        sectionName = Nothing

        If System.String.IsNullOrWhiteSpace(text) Then
            ShowCustomMessageBox("Import content is empty.")
            Return False
        End If

        ' Find first section header [ ... ]
        Dim lines As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserve(text)
        For Each line As System.String In lines
            Dim t As System.String = line.Trim()
            If t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso t.Length >= 2 Then
                sectionName = t.Substring(1, t.Length - 2).Trim()
                Exit For
            End If
        Next

        If System.String.IsNullOrWhiteSpace(sectionName) Then
            ' Ask user for a section name

            sectionName = ShowCustomInputBox($"You wish to import settings thar require a Section header (a user friendly name of the model or service, e.g., 'LexiSearch' or 'Gemini 3 Pro with minimum reasoning'). It can be changed later. Please enter a section name for the {friendlyType}:", TITLE_IMPORT, True, "Name")

            If System.String.IsNullOrWhiteSpace(sectionName) Then
                ShowCustomMessageBox("No section name provided. Import aborted.")
                Return False
            End If
            sectionName = sectionName.Trim()
        End If

        Return True

    End Function

    ' =========================================================================================
    '  DRY RUN PLAN + COMMIT
    ' =========================================================================================
    Private NotInheritable Class DryRunPlan
        Public Property TargetIniPath As System.String
        Public Property Kind As ImportKind
        Public Property TargetSectionName As System.String
        Public Property NewFileLines As System.Collections.Generic.List(Of System.String)
        Public Property RemovedLinesBackup As System.Collections.Generic.List(Of System.String)
        Public Property OverwrittenKeys As System.Collections.Generic.List(Of System.String)
        Public Property WillCreateRemovedBackup As System.Boolean

        Public Function GetUserSummary() As System.String
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("Dry run – review before importing")
            sb.AppendLine()
            sb.AppendLine("Target file: " & TargetIniPath)
            sb.AppendLine()

            If Kind = ImportKind.AlternateModel OrElse Kind = ImportKind.SpecialService Then
                sb.AppendLine("Target section: [" & TargetSectionName & "]")
                sb.AppendLine()
            End If

            If OverwrittenKeys IsNot Nothing AndAlso OverwrittenKeys.Count > 0 Then
                sb.AppendLine("Keys that will be overwritten (" & OverwrittenKeys.Count.ToString() & "):")
                For Each k As System.String In OverwrittenKeys
                    sb.AppendLine("  - " & k)
                Next
                sb.AppendLine()
            Else
                If Kind = ImportKind.PrimaryModel OrElse Kind = ImportKind.SecondaryModel OrElse Kind = ImportKind.OtherParameters Then
                    sb.AppendLine("No existing keys will be overwritten (new keys will be appended at the end of the file).")
                    sb.AppendLine()
                End If
            End If

            If WillCreateRemovedBackup Then
                sb.AppendLine("A full backup of the target file and a backup of the removed content will always be created in the same directory.")
            Else
                sb.AppendLine("A full backup of the target file will always be created in the same directory.")
            End If
            sb.AppendLine()
            sb.AppendLine("Proceed with import?")

            Return sb.ToString()
        End Function
    End Class

    Private Shared Function BuildDryRunPlan(context As ISharedContext,
                                           kind As ImportKind,
                                           targetIniPath As System.String,
                                           targetSectionName As System.String,
                                           importedLines As System.Collections.Generic.List(Of System.String)) As DryRunPlan

        If System.String.IsNullOrWhiteSpace(targetIniPath) Then Throw New System.ArgumentNullException(NameOf(targetIniPath))
        If importedLines Is Nothing OrElse importedLines.Count = 0 Then Throw New System.Exception("No imported lines.")

        Dim existingLines As System.Collections.Generic.List(Of System.String) = ReadAllLinesPreserve(targetIniPath)

        Dim plan As New DryRunPlan() With {
            .TargetIniPath = targetIniPath,
            .Kind = kind,
            .TargetSectionName = targetSectionName,
            .NewFileLines = New System.Collections.Generic.List(Of System.String)(existingLines),
            .RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)(),
            .OverwrittenKeys = New System.Collections.Generic.List(Of System.String)(),
            .WillCreateRemovedBackup = False
        }

        If kind = ImportKind.AlternateModel OrElse kind = ImportKind.SpecialService Then
            ' Replace entire section (Option A)
            Dim newSectionLines As System.Collections.Generic.List(Of System.String) = BuildSectionLines(targetSectionName, importedLines)
            ApplySectionReplace(plan, existingLines, newSectionLines)
            Return plan
        End If

        ' redink.ini modifications: parse imported key/value lines; apply secondary model logic if needed
        Dim kv As System.Collections.Generic.Dictionary(Of System.String, System.String) = ParseKeyValueLines(importedLines)

        If kind = ImportKind.SecondaryModel Then
            kv = ConvertKeysToSecondary(kv)
            ' ensure SecondAPI=True
            If Not kv.ContainsKey("SecondAPI") Then
                kv.Add("SecondAPI", "True")
            Else
                kv("SecondAPI") = "True"
            End If
        End If

        ApplyMainIniKeyReplaceAppend(plan, existingLines, kv)

        Return plan

    End Function

    Private Shared Function BuildDryRunPlanForSingleKey(mainIniPath As System.String, key As System.String, value As System.String) As DryRunPlan
        Dim existingLines As System.Collections.Generic.List(Of System.String) = ReadAllLinesPreserve(mainIniPath)

        Dim kv As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        kv(key) = value

        Dim plan As New DryRunPlan() With {
            .TargetIniPath = mainIniPath,
            .Kind = ImportKind.OtherParameters,
            .TargetSectionName = Nothing,
            .NewFileLines = New System.Collections.Generic.List(Of System.String)(existingLines),
            .RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)(),
            .OverwrittenKeys = New System.Collections.Generic.List(Of System.String)(),
            .WillCreateRemovedBackup = False
        }

        ApplyMainIniKeyReplaceAppend(plan, existingLines, kv)
        Return plan
    End Function

    Private Shared Sub CommitDryRunPlan(plan As DryRunPlan)

        If plan Is Nothing Then Throw New System.ArgumentNullException(NameOf(plan))
        If System.String.IsNullOrWhiteSpace(plan.TargetIniPath) Then Throw New System.Exception("Target ini path missing.")
        If plan.NewFileLines Is Nothing Then Throw New System.Exception("No new content.")

        Dim targetPath As System.String = plan.TargetIniPath

        Dim targetDir As System.String = System.IO.Path.GetDirectoryName(targetPath)
        If System.String.IsNullOrWhiteSpace(targetDir) Then
            Throw New System.Exception("Invalid target path.")
        End If

        If Not System.IO.Directory.Exists(targetDir) Then
            Throw New System.Exception("Target directory does not exist: " & targetDir)
        End If

        ' ------------------------------------------------------------------
        ' Full backup first (NO locking here)
        ' ------------------------------------------------------------------
        Dim ts As System.String = System.DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim baseName As System.String = System.IO.Path.GetFileNameWithoutExtension(targetPath)
        Dim ext As System.String = System.IO.Path.GetExtension(targetPath)

        Dim fullBackup As System.String =
        System.IO.Path.Combine(targetDir, baseName & "_" & ts & ".bak")

        System.IO.File.Copy(targetPath, fullBackup, overwrite:=False)

        ' ------------------------------------------------------------------
        ' Removed-content backup (if applicable)
        ' ------------------------------------------------------------------
        If plan.WillCreateRemovedBackup AndAlso
       plan.RemovedLinesBackup IsNot Nothing AndAlso
       plan.RemovedLinesBackup.Count > 0 Then

            Dim removedBackup As System.String =
            System.IO.Path.Combine(targetDir, baseName & "_removed_" & ts & ".bak")

            System.IO.File.WriteAllText(
            removedBackup,
            System.String.Join(System.Environment.NewLine, plan.RemovedLinesBackup),
            System.Text.Encoding.UTF8
        )
        End If

        ' ------------------------------------------------------------------
        ' Write temp file
        ' ------------------------------------------------------------------
        Dim tmpPath As System.String =
        System.IO.Path.Combine(
            targetDir,
            baseName & "_tmp_" & System.Guid.NewGuid().ToString("N") & ext
        )

        System.IO.File.WriteAllText(
        tmpPath,
        System.String.Join(System.Environment.NewLine, plan.NewFileLines),
        System.Text.Encoding.UTF8
    )

        ' ------------------------------------------------------------------
        ' Atomic replace (Windows handles locking)
        ' ------------------------------------------------------------------
        Try
            System.IO.File.Replace(tmpPath, targetPath, Nothing, True)
        Catch
            ' Fallback (still safe because temp is on same volume)
            Try
                System.IO.File.Delete(targetPath)
            Catch
                ' ignore
            End Try
            System.IO.File.Move(tmpPath, targetPath)
        End Try

    End Sub



    ' =========================================================================================
    '  APPLY: redink.ini key overwrite (remove old lines, append new at end)
    ' =========================================================================================
    Private Shared Sub ApplyMainIniKeyReplaceAppend(plan As DryRunPlan,
                                                   existingLines As System.Collections.Generic.List(Of System.String),
                                                   newKeyValues As System.Collections.Generic.Dictionary(Of System.String, System.String))

        Dim keys As New System.Collections.Generic.HashSet(Of System.String)(newKeyValues.Keys, System.StringComparer.OrdinalIgnoreCase)

        Dim newLines As New System.Collections.Generic.List(Of System.String)()
        Dim removed As New System.Collections.Generic.List(Of System.String)()
        Dim overwritten As New System.Collections.Generic.List(Of System.String)()

        For Each line As System.String In existingLines
            Dim parsedKey As System.String = Nothing
            Dim isKeyLine As System.Boolean = TryParseIniKey(line, parsedKey)

            If isKeyLine AndAlso Not System.String.IsNullOrWhiteSpace(parsedKey) AndAlso keys.Contains(parsedKey) Then
                removed.Add(line)
                overwritten.Add(parsedKey)
                Continue For ' remove it
            End If

            newLines.Add(line)
        Next

        ' Append a blank line before appended block if file doesn't end with blank line
        If newLines.Count > 0 Then
            Dim last As System.String = newLines(newLines.Count - 1)
            If last IsNot Nothing AndAlso last.Trim().Length > 0 Then
                newLines.Add("")
            End If
        End If

        ' Append new key/value lines at end (keep exactly "key = value" style)
        For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In newKeyValues
            newLines.Add(kvp.Key & " = " & kvp.Value)
        Next

        plan.NewFileLines = newLines

        If removed.Count > 0 Then
            plan.WillCreateRemovedBackup = True
            plan.RemovedLinesBackup = removed
            plan.OverwrittenKeys = UniquePreserveOrder(overwritten)
        Else
            plan.WillCreateRemovedBackup = False
            plan.RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)()
            plan.OverwrittenKeys = New System.Collections.Generic.List(Of System.String)()
        End If

    End Sub

    Private Shared Function UniquePreserveOrder(items As System.Collections.Generic.List(Of System.String)) As System.Collections.Generic.List(Of System.String)
        Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Dim res As New System.Collections.Generic.List(Of System.String)()
        For Each s As System.String In items
            If System.String.IsNullOrWhiteSpace(s) Then Continue For
            If seen.Add(s) Then res.Add(s)
        Next
        Return res
    End Function

    ' =========================================================================================
    '  APPLY: section replace (Option A)
    ' =========================================================================================
    Private Shared Sub ApplySectionReplace(plan As DryRunPlan,
                                          existingLines As System.Collections.Generic.List(Of System.String),
                                          newSectionLines As System.Collections.Generic.List(Of System.String))

        Dim sectionName As System.String = plan.TargetSectionName
        If System.String.IsNullOrWhiteSpace(sectionName) Then Throw New System.Exception("Section name missing.")

        Dim startIndex As System.Int32 = -1
        Dim endIndex As System.Int32 = -1

        FindSectionRange(existingLines, sectionName, startIndex, endIndex)

        Dim newFile As New System.Collections.Generic.List(Of System.String)()

        If startIndex >= 0 AndAlso endIndex >= startIndex Then
            ' Section exists -> replace in place, preserve order relevance
            ' Removed backup: include the exact lines being removed
            Dim removed As New System.Collections.Generic.List(Of System.String)()
            For i As System.Int32 = startIndex To endIndex
                removed.Add(existingLines(i))
            Next

            ' Copy before section
            For i As System.Int32 = 0 To startIndex - 1
                newFile.Add(existingLines(i))
            Next

            ' Insert new section
            For Each l As System.String In newSectionLines
                newFile.Add(l)
            Next

            ' Copy after section
            For i As System.Int32 = endIndex + 1 To existingLines.Count - 1
                newFile.Add(existingLines(i))
            Next

            plan.NewFileLines = newFile
            plan.WillCreateRemovedBackup = True
            plan.RemovedLinesBackup = removed
            Dim overwrittenKeys As New System.Collections.Generic.List(Of System.String)()

            For Each line As String In removed
                Dim key As String = Nothing
                If TryParseIniKey(line, key) Then
                    overwrittenKeys.Add(key)
                End If
            Next

            plan.OverwrittenKeys = UniquePreserveOrder(overwrittenKeys)

        Else
            ' Section not found -> append at end
            newFile = New System.Collections.Generic.List(Of System.String)(existingLines)

            If newFile.Count > 0 Then
                Dim last As System.String = newFile(newFile.Count - 1)
                If last IsNot Nothing AndAlso last.Trim().Length > 0 Then
                    newFile.Add("")
                End If
            End If

            For Each l As System.String In newSectionLines
                newFile.Add(l)
            Next

            plan.NewFileLines = newFile
            plan.WillCreateRemovedBackup = False
            plan.RemovedLinesBackup = New System.Collections.Generic.List(Of System.String)()
            plan.OverwrittenKeys = New System.Collections.Generic.List(Of System.String)()
        End If

    End Sub

    Private Shared Sub FindSectionRange(lines As System.Collections.Generic.List(Of System.String),
                                        sectionName As System.String,
                                        ByRef startIndex As System.Int32,
                                        ByRef endIndex As System.Int32)

        startIndex = -1
        endIndex = -1
        If lines Is Nothing OrElse lines.Count = 0 Then Return

        Dim targetHeader As System.String = "[" & sectionName & "]"

        For i As System.Int32 = 0 To lines.Count - 1
            Dim t As System.String = lines(i)
            If t Is Nothing Then Continue For
            Dim trimmed As System.String = t.Trim()
            If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") Then
                If System.String.Equals(trimmed, targetHeader, System.StringComparison.OrdinalIgnoreCase) Then
                    startIndex = i
                    Exit For
                End If
            End If
        Next

        If startIndex < 0 Then Return

        ' endIndex is last line before next section header (or EOF)
        endIndex = lines.Count - 1
        For i As System.Int32 = startIndex + 1 To lines.Count - 1
            Dim trimmed As System.String = If(lines(i), "").Trim()
            If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") AndAlso trimmed.Length >= 2 Then
                endIndex = i - 1
                Exit For
            End If
        Next

    End Sub

    Private Shared Function BuildSectionLines(sectionName As System.String,
                                              importedLines As System.Collections.Generic.List(Of System.String)) As System.Collections.Generic.List(Of System.String)

        Dim res As New System.Collections.Generic.List(Of System.String)()
        res.Add("[" & sectionName & "]")
        res.Add("")

        ' Remove any leading section headers in importedLines (we will write our own)
        For Each line As System.String In importedLines
            Dim t As System.String = line.Trim()
            If t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso t.Length >= 2 Then
                Continue For
            End If
            res.Add(line)
        Next

        Return res
    End Function

    ' =========================================================================================
    '  PARSING UTILITIES
    ' =========================================================================================
    Private Shared Function SplitToLinesPreserve(text As System.String) As System.Collections.Generic.List(Of System.String)
        Dim res As New System.Collections.Generic.List(Of System.String)()
        If text Is Nothing Then Return res

        ' Normalize to \n then split
        Dim normalized As System.String = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim parts As System.String() = normalized.Split(New System.Char() {ControlChars.Lf}, System.StringSplitOptions.None)
        res.AddRange(parts)
        Return res
    End Function

    Private Shared Function SplitToLinesPreserveNonEmpty(text As System.String) As System.Collections.Generic.List(Of System.String)
        Dim all As System.Collections.Generic.List(Of System.String) = SplitToLinesPreserve(text)
        Dim res As New System.Collections.Generic.List(Of System.String)()
        For Each l As System.String In all
            If l Is Nothing Then Continue For
            ' keep comment/blank too? For imports we keep everything as-is; but for parsing kv we need kv lines.
            ' We'll keep all lines; later ParseKeyValueLines will ignore non-kv/comment as needed.
            res.Add(l)
        Next
        Return res
    End Function

    Private Shared Function ReadAllLinesPreserve(path As System.String) As System.Collections.Generic.List(Of System.String)
        Dim text As System.String = System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8)
        Return SplitToLinesPreserve(text)
    End Function

    Private Shared Function TryParseIniKey(line As System.String, ByRef key As System.String) As System.Boolean
        key = Nothing
        If line Is Nothing Then Return False

        Dim trimmed As System.String = line.TrimStart()
        If trimmed.StartsWith(";", System.StringComparison.Ordinal) Then Return False
        If trimmed.StartsWith("[", System.StringComparison.Ordinal) Then Return False

        Dim idx As System.Int32 = line.IndexOf("="c)
        If idx <= 0 Then Return False

        Dim left As System.String = line.Substring(0, idx).Trim()
        If System.String.IsNullOrWhiteSpace(left) Then Return False

        key = left
        Return True
    End Function

    Private Shared Function ParseKeyValueLines(lines As System.Collections.Generic.List(Of System.String)) As System.Collections.Generic.Dictionary(Of System.String, System.String)
        Dim kv As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)

        For Each line As System.String In lines
            If line Is Nothing Then Continue For

            Dim trimmedStart As System.String = line.TrimStart()
            If trimmedStart.StartsWith(";", System.StringComparison.Ordinal) Then Continue For
            If trimmedStart.StartsWith("[", System.StringComparison.Ordinal) Then Continue For

            Dim idx As System.Int32 = line.IndexOf("="c)
            If idx <= 0 Then Continue For

            Dim k As System.String = line.Substring(0, idx).Trim()
            Dim v As System.String = line.Substring(idx + 1).Trim()

            If System.String.IsNullOrWhiteSpace(k) Then Continue For

            ' Keep last occurrence from import text (deterministic)
            kv(k) = v
        Next

        Return kv
    End Function


    ' =========================================================================================
    '  MULTI-SEGMENT PARSER (SectionName -> Lines)
    ' =========================================================================================
    Private Shared Function ParseIniSegments(
    lines As System.Collections.Generic.List(Of System.String)
) As System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.List(Of System.String))

        Dim result As New System.Collections.Generic.Dictionary(
        Of System.String,
        System.Collections.Generic.List(Of System.String)
    )(System.StringComparer.OrdinalIgnoreCase)

        Dim currentSection As System.String = Nothing
        Dim currentLines As System.Collections.Generic.List(Of System.String) = Nothing

        For Each line As System.String In lines
            Dim t As System.String = If(line, "").Trim()

            If t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso t.Length > 2 Then
                currentSection = t.Substring(1, t.Length - 2).Trim()
                currentLines = New System.Collections.Generic.List(Of System.String)()
                result(currentSection) = currentLines
            ElseIf currentSection IsNot Nothing Then
                currentLines.Add(line)
            End If
        Next

        Return result
    End Function


    Private Shared Function ConvertKeysToSecondary(kv As System.Collections.Generic.Dictionary(Of System.String, System.String)) As System.Collections.Generic.Dictionary(Of System.String, System.String)
        Dim out As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)

        For Each kvp As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In kv
            Dim k As System.String = kvp.Key
            Dim v As System.String = kvp.Value

            If System.String.IsNullOrWhiteSpace(k) Then Continue For

            If System.String.Equals(k, "SecondAPI", System.StringComparison.OrdinalIgnoreCase) Then
                out("SecondAPI") = "True"
            ElseIf k.EndsWith("_2", System.StringComparison.OrdinalIgnoreCase) Then
                out(k) = v
            Else
                out(k & "_2") = v
            End If
        Next

        Return out
    End Function

End Class
