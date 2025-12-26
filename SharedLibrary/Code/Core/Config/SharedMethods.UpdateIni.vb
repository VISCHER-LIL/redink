' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.UpdateIni.vb
' Purpose: Provides automatic INI configuration file updates from local or remote
'          sources with digital signature verification and user approval workflow.
'
' Architecture:
'  - Update Sources: Supports local file paths, network shares, and HTTPS URLs
'  - Digital Signatures: Ed25519 signatures via BouncyCastle for authenticity verification
'  - User Approval: Shows changes in a dialog allowing per-parameter approval/rejection
'  - Ignore List: Persists rejected parameters in My.Settings for future exclusion
'  - Three INI Files: redink.ini (main), allmodels.ini (alternate models), specialservices.ini
'
' Configuration Keys (to be added to ISharedContext):
'  - INI_UpdateIni: Boolean - Master switch for update mechanism
'  - INI_UpdateIniAllowRemote: Boolean - Allow HTTPS sources (vs local/network only)
'  - INI_UpdateIniNoSignature: Boolean - Skip signature verification if True
'  - INI_UpdateSource: String - Update source for redink.ini (path; keys; pubkey)
'
' My.Settings Variables:
'  - IgnoredUpdates_RedInk: String - Serialized ignored parameters for redink.ini
'  - IgnoredUpdates_AlternateModels: String - Serialized ignored params per segment
'  - IgnoredUpdates_SpecialServices: String - Serialized ignored params per segment
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Crypto.Generators
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Crypto.Signers
Imports Org.BouncyCastle.Security
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

#Region "Temporary Context Variable Assignments"
        ' =============================================================================
        ' TEMPORARY: These properties should be moved to ISharedContext and loaded
        ' in InitializeConfig from the INI file. For now, they are placeholders.
        ' =============================================================================

        ' Master switch: If False, no update checks are performed
        Private Shared _INI_UpdateIni As Boolean = False
        Public Shared Property INI_UpdateIni As Boolean
            Get
                Return _INI_UpdateIni
            End Get
            Set(value As Boolean)
                _INI_UpdateIni = value
            End Set
        End Property

        ' If False, only local file paths and network shares are allowed (no HTTPS)
        Private Shared _INI_UpdateIniAllowRemote As Boolean = True
        Public Shared Property INI_UpdateIniAllowRemote As Boolean
            Get
                Return _INI_UpdateIniAllowRemote
            End Get
            Set(value As Boolean)
                _INI_UpdateIniAllowRemote = value
            End Set
        End Property

        ' If True, signature verification is skipped (NOT RECOMMENDED for production)
        Private Shared _INI_UpdateIniNoSignature As Boolean = False
        Public Shared Property INI_UpdateIniNoSignature As Boolean
            Get
                Return _INI_UpdateIniNoSignature
            End Get
            Set(value As Boolean)
                _INI_UpdateIniNoSignature = value
            End Set
        End Property

        ' Update source for redink.ini: "path; keylist; base64_public_key"
        Private Shared _INI_UpdateSource As String = ""
        Public Shared Property INI_UpdateSource As String
            Get
                Return _INI_UpdateSource
            End Get
            Set(value As String)
                _INI_UpdateSource = value
            End Set
        End Property

#End Region

#Region "Data Structures"

        ''' <summary>
        ''' Represents a single parameter change detected during update check.
        ''' </summary>
        Public Class IniParameterChange
            Public Property IniFile As String           ' "redink.ini", "allmodels.ini", "specialservices.ini"
            Public Property SegmentName As String       ' Empty for redink.ini, segment name for others
            Public Property ParameterKey As String      ' The key name (e.g., "Model", "Endpoint")
            Public Property OldValue As String          ' Current value in local file
            Public Property NewValue As String          ' Value from remote source
            Public Property IsSelected As Boolean       ' User's approval choice
            Public Property IsSuspicious As Boolean     ' True if contains URL/path that changed

            Public Overrides Function ToString() As String
                Return $"[{IniFile}]{If(String.IsNullOrEmpty(SegmentName), "", $"[{SegmentName}]")}.{ParameterKey}"
            End Function
        End Class

        ''' <summary>
        ''' Represents an INI file segment with its parameters and update source.
        ''' </summary>
        Public Class IniSegment
            Public Property Name As String                                      ' Segment name (from [Name])
            Public Property Parameters As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Public Property UpdateSource As String                              ' Full UpdateSource value
            Public Property UpdatePath As String                                ' Extracted path/URL
            Public Property UpdateKeys As List(Of String)                       ' Keys to update ("all" or specific list)
            Public Property PublicKey As String                                 ' Base64 public key for signature
        End Class

        ''' <summary>
        ''' Represents a parsed INI file with optional segments.
        ''' </summary>
        Public Class ParsedIniFile
            Public Property FilePath As String
            Public Property FileName As String
            Public Property GlobalParameters As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Public Property Segments As New List(Of IniSegment)()
            Public Property RawContent As String
        End Class

#End Region

#Region "Keys That Cannot Be Updated"

        ''' <summary>
        ''' Returns True if the specified key should never be updated automatically.
        ''' </summary>
        Private Shared Function IsProtectedKey(key As String) As Boolean
            If String.IsNullOrWhiteSpace(key) Then Return True
            ' Keys starting with "Update" are protected
            Return key.Trim().StartsWith("Update", StringComparison.OrdinalIgnoreCase)
        End Function

#End Region

#Region "Main Entry Point"

        ''' <summary>
        ''' Main entry point for INI update checking. Called from UpdateHandler at startup.
        ''' </summary>
        ''' <param name="context">The shared context containing configuration.</param>
        ''' <returns>True if updates were applied, False otherwise.</returns>
        Public Shared Function CheckForIniUpdates(ByRef context As ISharedContext) As Boolean
            Try
                ' Check master switch
                If Not INI_UpdateIni Then
                    Debug.WriteLine("INI Update: Disabled via INI_UpdateIni")
                    Return False
                End If

                ' Collect all changes from all three INI files
                Dim allChanges As New List(Of IniParameterChange)()
                Dim signatureErrors As New List(Of String)()

                ' 1. Check redink.ini (main config)
                Dim mainIniPath As String = GetDefaultINIPath(context.RDV)
                If File.Exists(mainIniPath) AndAlso Not String.IsNullOrWhiteSpace(INI_UpdateSource) Then
                    Dim mainFileName = Path.GetFileName(mainIniPath)
                    Dim mainChanges = CheckSingleIniFile(mainIniPath, mainFileName, INI_UpdateSource, Nothing, signatureErrors)
                    If mainChanges IsNot Nothing Then allChanges.AddRange(mainChanges)
                End If

                ' 2. Check AlternateModelPath (user-defined filename)
                If Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                    Dim altPath = ExpandEnvironmentVariables(context.INI_AlternateModelPath)
                    If File.Exists(altPath) Then
                        Dim altFileName = Path.GetFileName(altPath)
                        Dim altChanges = CheckSegmentedIniFile(altPath, altFileName, signatureErrors)
                        If altChanges IsNot Nothing Then allChanges.AddRange(altChanges)
                    End If
                End If

                ' 3. Check SpecialServicePath (user-defined filename)
                If Not String.IsNullOrWhiteSpace(context.INI_SpecialServicePath) Then
                    Dim svcPath = ExpandEnvironmentVariables(context.INI_SpecialServicePath)
                    If File.Exists(svcPath) Then
                        Dim svcFileName = Path.GetFileName(svcPath)
                        Dim svcChanges = CheckSegmentedIniFile(svcPath, svcFileName, signatureErrors)
                        If svcChanges IsNot Nothing Then allChanges.AddRange(svcChanges)
                    End If
                End If

                ' Report any signature errors to the user
                If signatureErrors.Count > 0 Then
                    ShowSignatureErrorDialog(signatureErrors)
                End If

                ' Filter out ignored parameters
                allChanges = FilterIgnoredParameters(allChanges)

                If allChanges.Count = 0 Then
                    Debug.WriteLine("INI Update: No changes detected")
                    Return False
                End If

                ' Show approval dialog
                Dim approvalResult = ShowUpdateApprovalDialog(allChanges)
                If approvalResult = UpdateApprovalResult.Reject Then
                    ' User rejected all - show ignore dialog
                    ShowIgnoreConfirmationDialog(allChanges)
                    Return False
                ElseIf approvalResult = UpdateApprovalResult.Cancel Then
                    Return False
                End If

                ' Get approved and rejected changes
                Dim approvedChanges = allChanges.Where(Function(c) c.IsSelected).ToList()
                Dim rejectedChanges = allChanges.Where(Function(c) Not c.IsSelected).ToList()

                ' Show ignore dialog for rejected items
                If rejectedChanges.Count > 0 Then
                    ShowIgnoreConfirmationDialog(rejectedChanges)
                End If

                ' Apply approved changes
                If approvedChanges.Count > 0 Then
                    ApplyApprovedUpdates(approvedChanges, context)
                    ShowCustomMessageBox($"{approvedChanges.Count} configuration parameter(s) have been updated. Changes will be active upon next reload.")
                    Return True
                End If

                Return False

            Catch ex As Exception
                Debug.WriteLine($"INI Update Error: {ex.Message}")
                Return False
            End Try
        End Function

        Private Enum UpdateApprovalResult
            Approve
            Reject
            Cancel
        End Enum

#End Region

#Region "Signature Error Reporting"

        ''' <summary>
        ''' Represents a signature validation error for reporting.
        ''' </summary>
        Public Class SignatureError
            Public Property SourcePath As String
            Public Property ErrorType As SignatureErrorType
            Public Property Details As String

            Public Overrides Function ToString() As String
                Return $"[{ErrorType}] {SourcePath}: {Details}"
            End Function
        End Class

        Public Enum SignatureErrorType
            SignatureFileMissing
            PublicKeyMissing
            SignatureInvalid
            DownloadFailed
            Other
        End Enum

        ''' <summary>
        ''' Shows a dialog to the user reporting signature validation errors.
        ''' This alerts users to potential security issues or configuration problems.
        ''' </summary>
        Private Shared Sub ShowSignatureErrorDialog(errors As List(Of String))
            If errors Is Nothing OrElse errors.Count = 0 Then Return

            Dim form As New Form() With {
                .Text = $"{AN} - Signature Validation Errors",
                .Size = New Size(750, 450),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(500, 300)
            }

            Try
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Warning header with icon
            Dim pnlHeader As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 70,
                .BackColor = Color.FromArgb(255, 250, 230)
            }
            form.Controls.Add(pnlHeader)

            Dim lblWarning As New Label() With {
                .Text = "⚠ SECURITY WARNING: The following update sources could not be verified." & vbCrLf &
                        "This may indicate tampering, misconfiguration, or missing signature files.",
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
                .ForeColor = Color.DarkOrange,
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(15, 10, 15, 10)
            }
            pnlHeader.Controls.Add(lblWarning)

            ' Error list
            Dim txtErrors As New TextBox() With {
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Both,
                .BackColor = SystemColors.Window,
                .Font = New Font("Consolas", 9.0F),
                .Text = String.Join(vbCrLf & vbCrLf, errors)
            }
            form.Controls.Add(txtErrors)

            ' Info panel
            Dim lblInfo As New Label() With {
                .Text = "These update sources were skipped. Contact your administrator if this is unexpected." & vbCrLf &
                        "Administrators: Use the Signature Management tool to diagnose and fix signature issues.",
                .Dock = DockStyle.Bottom,
                .Height = 50,
                .Padding = New Padding(10),
                .BackColor = Color.FromArgb(240, 240, 240)
            }
            form.Controls.Add(lblInfo)

            ' Buttons
            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .Height = 50,
                .Padding = New Padding(10)
            }
            form.Controls.Add(pnlButtons)

            Dim btnClose As New Button() With {.Text = "Close", .AutoSize = True}
            Dim btnCopy As New Button() With {.Text = "Copy to Clipboard", .AutoSize = True}
            Dim btnDiagnose As New Button() With {.Text = "Open Signature Tool...", .AutoSize = True}

            pnlButtons.Controls.Add(btnClose)
            pnlButtons.Controls.Add(btnCopy)
            pnlButtons.Controls.Add(btnDiagnose)

            AddHandler btnClose.Click, Sub() form.Close()

            AddHandler btnCopy.Click, Sub()
                                          Try
                                              Dim report = $"=== {AN} Signature Validation Report ==={vbCrLf}" &
                                                           $"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}{vbCrLf}{vbCrLf}" &
                                                           String.Join(vbCrLf & vbCrLf, errors)
                                              Clipboard.SetText(report)
                                              ShowCustomMessageBox("Error report copied to clipboard.")
                                          Catch ex As Exception
                                              ShowCustomMessageBox("Failed to copy to clipboard: " & ex.Message)
                                          End Try
                                      End Sub

            AddHandler btnDiagnose.Click, Sub()
                                              form.Close()
                                              ShowSignatureManagementDialog()
                                          End Sub

            txtErrors.BringToFront()
            form.ShowDialog()
        End Sub

#End Region

#Region "INI File Parsing"

        ''' <summary>
        ''' Parses an INI file into a structured representation with segments.
        ''' </summary>
        Private Shared Function ParseIniFile(filePath As String) As ParsedIniFile
            Dim result As New ParsedIniFile() With {
                .FilePath = filePath,
                .FileName = Path.GetFileName(filePath)
            }

            If Not File.Exists(filePath) Then Return result

            Try
                result.RawContent = File.ReadAllText(filePath, Encoding.UTF8)
                Dim lines = result.RawContent.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)

                Dim currentSegment As IniSegment = Nothing

                For Each rawLine In lines
                    Dim line = rawLine.Trim()

                    ' Skip empty lines and comments
                    If String.IsNullOrEmpty(line) OrElse line.StartsWith(";") Then Continue For

                    ' Check for segment header [Name]
                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        ' Save previous segment if exists
                        If currentSegment IsNot Nothing Then
                            ParseUpdateSource(currentSegment)
                            result.Segments.Add(currentSegment)
                        End If

                        ' Start new segment
                        currentSegment = New IniSegment() With {
                            .Name = line.Substring(1, line.Length - 2).Trim()
                        }
                        Continue For
                    End If

                    ' Parse key = value
                    Dim eqIndex = line.IndexOf("="c)
                    If eqIndex > 0 Then
                        Dim key = line.Substring(0, eqIndex).Trim()
                        Dim value = line.Substring(eqIndex + 1).Trim()

                        If currentSegment IsNot Nothing Then
                            currentSegment.Parameters(key) = value
                        Else
                            result.GlobalParameters(key) = value
                        End If
                    End If
                Next

                ' Save last segment
                If currentSegment IsNot Nothing Then
                    ParseUpdateSource(currentSegment)
                    result.Segments.Add(currentSegment)
                End If

            Catch ex As Exception
                Debug.WriteLine($"Error parsing INI file {filePath}: {ex.Message}")
            End Try

            Return result
        End Function

        ''' <summary>
        ''' Parses the UpdateSource parameter of a segment into its components.
        ''' Format: "path; key1,key2,key3; base64_public_key" or "path; all; base64_public_key"
        ''' </summary>
        Private Shared Sub ParseUpdateSource(segment As IniSegment)
            If segment Is Nothing Then Return
            If Not segment.Parameters.ContainsKey("UpdateSource") Then Return

            segment.UpdateSource = segment.Parameters("UpdateSource")
            Dim parts = segment.UpdateSource.Split(";"c)

            If parts.Length >= 1 Then
                segment.UpdatePath = parts(0).Trim()
            End If

            If parts.Length >= 2 Then
                Dim keysPart = parts(1).Trim()
                If keysPart.Equals("all", StringComparison.OrdinalIgnoreCase) Then
                    segment.UpdateKeys = New List(Of String) From {"*"} ' Special marker for "all"
                Else
                    segment.UpdateKeys = keysPart.Split(","c).
                        Select(Function(k) k.Trim()).
                        Where(Function(k) Not String.IsNullOrEmpty(k)).
                        ToList()
                End If
            End If

            If parts.Length >= 3 Then
                segment.PublicKey = parts(2).Trim()
            End If
        End Sub

        ''' <summary>
        ''' Parses UpdateSource for a global (non-segmented) INI file.
        ''' </summary>
        Private Shared Function ParseGlobalUpdateSource(updateSource As String) As IniSegment
            Dim segment As New IniSegment() With {
                .Name = "",
                .UpdateSource = updateSource
            }

            If String.IsNullOrWhiteSpace(updateSource) Then Return segment

            Dim parts = updateSource.Split(";"c)

            If parts.Length >= 1 Then
                segment.UpdatePath = parts(0).Trim()
            End If

            If parts.Length >= 2 Then
                Dim keysPart = parts(1).Trim()
                If keysPart.Equals("all", StringComparison.OrdinalIgnoreCase) Then
                    segment.UpdateKeys = New List(Of String) From {"*"}
                Else
                    segment.UpdateKeys = keysPart.Split(","c).
                        Select(Function(k) k.Trim()).
                        Where(Function(k) Not String.IsNullOrEmpty(k)).
                        ToList()
                End If
            End If

            If parts.Length >= 3 Then
                segment.PublicKey = parts(2).Trim()
            End If

            Return segment
        End Function

#End Region

#Region "Update Source Loading"

        ''' <summary>
        ''' Loads content from an update source (local, network, or HTTPS).
        ''' </summary>
        ''' <param name="sourcePath">Path or URL to the update source.</param>
        ''' <param name="publicKey">Base64-encoded public key for signature verification.</param>
        ''' <param name="signatureErrors">List to collect signature error messages for user reporting.</param>
        ''' <returns>Content string if successful, Nothing if failed.</returns>
        Private Shared Function LoadUpdateSourceContent(sourcePath As String, publicKey As String, signatureErrors As List(Of String)) As String
            If String.IsNullOrWhiteSpace(sourcePath) Then Return Nothing

            Try
                Dim isRemote = sourcePath.StartsWith("https://", StringComparison.OrdinalIgnoreCase) OrElse
                               sourcePath.StartsWith("http://", StringComparison.OrdinalIgnoreCase)

                ' Check if remote sources are allowed
                If isRemote AndAlso Not INI_UpdateIniAllowRemote Then
                    Debug.WriteLine($"Remote update source blocked by policy: {sourcePath}")
                    Return Nothing
                End If

                Dim content As String = Nothing
                Dim signatureContent As String = Nothing
                Dim expandedPath As String = Nothing

                If isRemote Then
                    ' Enable TLS 1.2 for HTTPS
                    ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12

                    Using client As New HttpClient()
                        client.Timeout = TimeSpan.FromSeconds(30)

                        ' Download main content
                        Try
                            Dim contentTask = client.GetStringAsync(sourcePath)
                            contentTask.Wait()
                            content = contentTask.Result
                        Catch ex As Exception
                            signatureErrors?.Add($"SOURCE: {sourcePath}" & vbCrLf &
                                                 $"ERROR: Failed to download update source" & vbCrLf &
                                                 $"DETAILS: {ex.Message}")
                            Return Nothing
                        End Try

                        ' Download signature file (.sig)
                        If Not INI_UpdateIniNoSignature Then
                            Try
                                Dim sigTask = client.GetStringAsync(sourcePath & ".sig")
                                sigTask.Wait()
                                signatureContent = sigTask.Result?.Trim()
                            Catch ex As Exception
                                signatureErrors?.Add($"SOURCE: {sourcePath}" & vbCrLf &
                                                     $"ERROR: Signature file not found or inaccessible" & vbCrLf &
                                                     $"EXPECTED: {sourcePath}.sig" & vbCrLf &
                                                     $"DETAILS: {ex.Message}" & vbCrLf &
                                                     $"ACTION: Ensure the .sig file exists alongside the update file, or contact your administrator.")
                            End Try
                        End If
                    End Using
                Else
                    ' Local or network path - expand environment variables
                    expandedPath = ExpandEnvironmentVariables(sourcePath)
                    If Not File.Exists(expandedPath) Then
                        Debug.WriteLine($"Update source file not found: {expandedPath}")
                        Return Nothing
                    End If

                    Try
                        content = File.ReadAllText(expandedPath, Encoding.UTF8)
                    Catch ex As Exception
                        signatureErrors?.Add($"SOURCE: {expandedPath}" & vbCrLf &
                                             $"ERROR: Failed to read update source file" & vbCrLf &
                                             $"DETAILS: {ex.Message}")
                        Return Nothing
                    End Try

                    ' Check for signature file
                    If Not INI_UpdateIniNoSignature Then
                        Dim sigPath = expandedPath & ".sig"
                        If File.Exists(sigPath) Then
                            Try
                                signatureContent = File.ReadAllText(sigPath, Encoding.UTF8)?.Trim()
                            Catch ex As Exception
                                signatureErrors?.Add($"SOURCE: {expandedPath}" & vbCrLf &
                                                     $"ERROR: Failed to read signature file" & vbCrLf &
                                                     $"SIGNATURE FILE: {sigPath}" & vbCrLf &
                                                     $"DETAILS: {ex.Message}")
                            End Try
                        Else
                            signatureErrors?.Add($"SOURCE: {expandedPath}" & vbCrLf &
                                                 $"ERROR: Signature file not found" & vbCrLf &
                                                 $"EXPECTED: {sigPath}" & vbCrLf &
                                                 $"ACTION: Create a .sig file using the Signature Management tool, or contact your administrator.")
                        End If
                    End If
                End If

                ' Verify signature if required
                If Not INI_UpdateIniNoSignature Then
                    Dim displayPath = If(expandedPath, sourcePath)

                    If String.IsNullOrWhiteSpace(signatureContent) Then
                        ' Error already added above
                        Return Nothing
                    End If

                    If String.IsNullOrWhiteSpace(publicKey) Then
                        signatureErrors?.Add($"SOURCE: {displayPath}" & vbCrLf &
                                             $"ERROR: No public key configured for signature verification" & vbCrLf &
                                             $"ACTION: Add the public key as the third parameter in UpdateSource:" & vbCrLf &
                                             $"        UpdateSource = path; keys; PUBLIC_KEY_HERE")
                        Return Nothing
                    End If

                    If Not VerifyEd25519Signature(content, signatureContent, publicKey) Then
                        signatureErrors?.Add($"SOURCE: {displayPath}" & vbCrLf &
                                             $"ERROR: SIGNATURE VERIFICATION FAILED" & vbCrLf &
                                             $"⚠ This may indicate the file has been tampered with!" & vbCrLf &
                                             $"POSSIBLE CAUSES:" & vbCrLf &
                                             $"  - File was modified after signing" & vbCrLf &
                                             $"  - Wrong public key configured" & vbCrLf &
                                             $"  - Signature file corrupted or for different file" & vbCrLf &
                                             $"ACTION: Contact your administrator immediately.")
                        Return Nothing
                    End If
                End If

                Return content

            Catch ex As Exception
                signatureErrors?.Add($"SOURCE: {sourcePath}" & vbCrLf &
                                     $"ERROR: Unexpected error during update check" & vbCrLf &
                                     $"DETAILS: {ex.Message}")
                Return Nothing
            End Try
        End Function

#End Region

#Region "Change Detection"

        ''' <summary>
        ''' Checks a single (non-segmented) INI file like redink.ini for updates.
        ''' </summary>
        Private Shared Function CheckSingleIniFile(localPath As String, fileName As String,
                                                   updateSource As String, segmentName As String,
                                                   signatureErrors As List(Of String)) As List(Of IniParameterChange)
            Dim changes As New List(Of IniParameterChange)()

            Try
                ' Parse update source
                Dim sourceInfo = ParseGlobalUpdateSource(updateSource)
                If String.IsNullOrWhiteSpace(sourceInfo.UpdatePath) Then Return changes
                If sourceInfo.UpdateKeys Is Nothing OrElse sourceInfo.UpdateKeys.Count = 0 Then Return changes

                ' Load remote content (with signature validation)
                Dim remoteContent = LoadUpdateSourceContent(sourceInfo.UpdatePath, sourceInfo.PublicKey, signatureErrors)
                If String.IsNullOrWhiteSpace(remoteContent) Then Return changes

                ' Parse local and remote files
                Dim localIni = ParseIniFile(localPath)
                Dim remoteParams = ParseIniContentToDict(remoteContent)

                ' Determine which keys to check
                Dim keysToCheck As IEnumerable(Of String)
                If sourceInfo.UpdateKeys.Contains("*") Then
                    ' "all" - check all keys from remote that exist locally
                    keysToCheck = remoteParams.Keys
                Else
                    keysToCheck = sourceInfo.UpdateKeys
                End If

                ' Compare values
                For Each key In keysToCheck
                    ' Skip protected keys
                    If IsProtectedKey(key) Then Continue For

                    If Not remoteParams.ContainsKey(key) Then Continue For

                    Dim remoteValue = remoteParams(key)
                    Dim localValue As String = Nothing
                    localIni.GlobalParameters.TryGetValue(key, localValue)

                    ' Check if values differ
                    If Not String.Equals(localValue, remoteValue, StringComparison.Ordinal) Then
                        Dim change As New IniParameterChange() With {
                            .IniFile = fileName,
                            .SegmentName = If(segmentName, ""),
                            .ParameterKey = key,
                            .OldValue = If(localValue, "(not set)"),
                            .NewValue = remoteValue,
                            .IsSelected = True,
                            .IsSuspicious = IsPathOrUrlChange(localValue, remoteValue)
                        }

                        ' Suspicious changes are not selected by default
                        If change.IsSuspicious Then change.IsSelected = False

                        changes.Add(change)
                    End If
                Next

            Catch ex As Exception
                Debug.WriteLine($"Error checking INI file {localPath}: {ex.Message}")
            End Try

            Return changes
        End Function

        ''' <summary>
        ''' Checks a segmented INI file (user-defined filename) for updates.
        ''' </summary>
        Private Shared Function CheckSegmentedIniFile(localPath As String, fileName As String,
                                                      signatureErrors As List(Of String)) As List(Of IniParameterChange)
            Dim changes As New List(Of IniParameterChange)()

            Try
                Dim localIni = ParseIniFile(localPath)

                For Each segment In localIni.Segments
                    ' Skip segments without UpdateSource
                    If String.IsNullOrWhiteSpace(segment.UpdatePath) Then Continue For
                    If segment.UpdateKeys Is Nothing OrElse segment.UpdateKeys.Count = 0 Then Continue For

                    ' Load remote content for this segment (with signature validation)
                    Dim remoteContent = LoadUpdateSourceContent(segment.UpdatePath, segment.PublicKey, signatureErrors)
                    If String.IsNullOrWhiteSpace(remoteContent) Then Continue For

                    ' Parse remote content - look for matching segment
                    Dim remoteSegment = FindSegmentInContent(remoteContent, segment.Name)
                    If remoteSegment Is Nothing Then Continue For

                    ' Determine which keys to check
                    Dim keysToCheck As IEnumerable(Of String)
                    If segment.UpdateKeys.Contains("*") Then
                        keysToCheck = remoteSegment.Parameters.Keys
                    Else
                        keysToCheck = segment.UpdateKeys
                    End If

                    ' Compare values
                    For Each key In keysToCheck
                        If IsProtectedKey(key) Then Continue For
                        If Not remoteSegment.Parameters.ContainsKey(key) Then Continue For

                        Dim remoteValue = remoteSegment.Parameters(key)
                        Dim localValue As String = Nothing
                        segment.Parameters.TryGetValue(key, localValue)

                        If Not String.Equals(localValue, remoteValue, StringComparison.Ordinal) Then
                            Dim change As New IniParameterChange() With {
                                .IniFile = fileName,
                                .SegmentName = segment.Name,
                                .ParameterKey = key,
                                .OldValue = If(localValue, "(not set)"),
                                .NewValue = remoteValue,
                                .IsSelected = True,
                                .IsSuspicious = IsPathOrUrlChange(localValue, remoteValue)
                            }

                            If change.IsSuspicious Then change.IsSelected = False

                            changes.Add(change)
                        End If
                    Next
                Next

            Catch ex As Exception
                Debug.WriteLine($"Error checking segmented INI file {localPath}: {ex.Message}")
            End Try

            Return changes
        End Function

        ''' <summary>
        ''' Parses INI content string into a simple key-value dictionary (no segments).
        ''' </summary>
        Private Shared Function ParseIniContentToDict(content As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(content) Then Return result

            Dim lines = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            For Each rawLine In lines
                Dim line = rawLine.Trim()
                If String.IsNullOrEmpty(line) OrElse line.StartsWith(";") OrElse line.StartsWith("[") Then Continue For

                Dim eqIndex = line.IndexOf("="c)
                If eqIndex > 0 Then
                    Dim key = line.Substring(0, eqIndex).Trim()
                    Dim value = line.Substring(eqIndex + 1).Trim()
                    result(key) = value
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' Finds a specific segment by name in INI content.
        ''' </summary>
        Private Shared Function FindSegmentInContent(content As String, segmentName As String) As IniSegment
            If String.IsNullOrWhiteSpace(content) OrElse String.IsNullOrWhiteSpace(segmentName) Then Return Nothing

            Dim lines = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim inTargetSegment = False
            Dim segment As IniSegment = Nothing

            For Each rawLine In lines
                Dim line = rawLine.Trim()

                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    ' End previous segment if we found our target
                    If inTargetSegment AndAlso segment IsNot Nothing Then
                        Return segment
                    End If

                    Dim name = line.Substring(1, line.Length - 2).Trim()
                    If name.Equals(segmentName, StringComparison.OrdinalIgnoreCase) Then
                        inTargetSegment = True
                        segment = New IniSegment() With {.Name = name}
                    Else
                        inTargetSegment = False
                    End If
                    Continue For
                End If

                If inTargetSegment AndAlso segment IsNot Nothing Then
                    If String.IsNullOrEmpty(line) OrElse line.StartsWith(";") Then Continue For

                    Dim eqIndex = line.IndexOf("="c)
                    If eqIndex > 0 Then
                        Dim key = line.Substring(0, eqIndex).Trim()
                        Dim value = line.Substring(eqIndex + 1).Trim()
                        segment.Parameters(key) = value
                    End If
                End If
            Next

            Return segment
        End Function

        ''' <summary>
        ''' Determines if a value change involves URL or path changes (suspicious).
        ''' </summary>
        Private Shared Function IsPathOrUrlChange(oldValue As String, newValue As String) As Boolean
            If String.IsNullOrWhiteSpace(oldValue) OrElse String.IsNullOrWhiteSpace(newValue) Then Return False

            ' Check if either contains URL patterns
            Dim urlPatterns = {"http://", "https://", "ftp://", "file://"}
            Dim pathPatterns = {":\", ":\\", "/"}

            Dim oldHasUrl = urlPatterns.Any(Function(p) oldValue.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
            Dim newHasUrl = urlPatterns.Any(Function(p) newValue.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
            Dim oldHasPath = pathPatterns.Any(Function(p) oldValue.Contains(p))
            Dim newHasPath = pathPatterns.Any(Function(p) newValue.Contains(p))

            ' Suspicious if URL or path is present and values differ
            If (oldHasUrl OrElse newHasUrl OrElse oldHasPath OrElse newHasPath) Then
                Return Not String.Equals(oldValue, newValue, StringComparison.OrdinalIgnoreCase)
            End If

            Return False
        End Function

#End Region

#Region "Ignored Parameters Management"

        ''' <summary>
        ''' Filters out parameters that are in the ignore list.
        ''' </summary>
        Private Shared Function FilterIgnoredParameters(changes As List(Of IniParameterChange)) As List(Of IniParameterChange)
            Dim result As New List(Of IniParameterChange)()

            For Each change In changes
                Dim ignoreKey = GetIgnoreKey(change)
                Dim ignoreList = GetIgnoreListForFile(change.IniFile)

                If Not ignoreList.Contains(ignoreKey) Then
                    result.Add(change)
                End If
            Next

            Return result
        End Function


        ''' <summary>
        ''' Generates a unique key for identifying an ignored parameter.
        ''' Includes the filename to support user-defined INI filenames.
        ''' </summary>
        Private Shared Function GetIgnoreKey(change As IniParameterChange) As String
            If String.IsNullOrEmpty(change.SegmentName) Then
                Return $"{change.IniFile}|{change.ParameterKey}"
            Else
                Return $"{change.IniFile}|{change.SegmentName}|{change.ParameterKey}"
            End If
        End Function

        ''' <summary>
        ''' Gets the ignore list for a specific INI file from My.Settings.
        ''' Uses the actual filename (not hardcoded names) to support user-defined filenames.
        ''' </summary>
        Private Shared Function GetIgnoreListForFile(fileName As String) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim settingsValue As String = ""

            Try
                ' Normalize filename to lowercase for consistent storage key
                Dim normalizedName = fileName.ToLowerInvariant()

                ' Map known config types to their settings keys
                ' For the main INI, use "RedInk" setting
                ' For other files, derive a key from the filename
                Dim settingsKey As String

                If normalizedName.EndsWith("redink.ini") OrElse normalizedName.EndsWith($"{AN2.ToLowerInvariant()}.ini") Then
                    settingsKey = "IgnoredUpdates_RedInk"
                Else
                    ' For user-defined filenames, use a hash-based approach or store in a single collection
                    ' For simplicity, we'll use a combined setting with filename prefix
                    settingsKey = "IgnoredUpdates_Custom"
                End If

                settingsValue = CStr(My.Settings.Item(settingsKey))
            Catch
                ' Settings may not exist yet
            End Try

            If Not String.IsNullOrWhiteSpace(settingsValue) Then
                For Each item In settingsValue.Split(";"c)
                    If Not String.IsNullOrWhiteSpace(item) Then
                        ' For custom files, items are stored as "filename|segment|key" or "filename|key"
                        result.Add(item.Trim())
                    End If
                Next
            End If

            Return result
        End Function

        ''' <summary>
        ''' Saves the ignore list for a specific INI file to My.Settings.
        ''' </summary>
        Private Shared Sub SaveIgnoreListForFile(fileName As String, ignoreList As HashSet(Of String))
            Try
                Dim value = String.Join(";", ignoreList)
                Dim normalizedName = fileName.ToLowerInvariant()

                Dim settingsKey As String
                If normalizedName.EndsWith("redink.ini") OrElse normalizedName.EndsWith($"{AN2.ToLowerInvariant()}.ini") Then
                    settingsKey = "IgnoredUpdates_RedInk"
                Else
                    settingsKey = "IgnoredUpdates_Custom"
                End If

                My.Settings.Item(settingsKey) = value
                My.Settings.Save()
            Catch ex As Exception
                Debug.WriteLine($"Error saving ignore list: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Adds parameters to the ignore list.
        ''' </summary>
        Private Shared Sub AddToIgnoreList(changes As List(Of IniParameterChange))
            ' Group by file
            Dim byFile = changes.GroupBy(Function(c) c.IniFile)

            For Each fileGroup In byFile
                Dim ignoreList = GetIgnoreListForFile(fileGroup.Key)

                For Each change In fileGroup
                    ignoreList.Add(GetIgnoreKey(change))
                Next

                SaveIgnoreListForFile(fileGroup.Key, ignoreList)
            Next
        End Sub

        ''' <summary>
        ''' Shows a dialog to manage ignored parameters.
        ''' </summary>
        Public Shared Sub ShowIgnoredParametersDialog()
            Dim form As New Form() With {
                .Text = $"{AN} - Manage Ignored Update Parameters",
                .Size = New Size(700, 500),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F)
            }

            Try
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Dim lblInfo As New Label() With {
                .Text = "The following parameters are ignored during automatic updates. Uncheck items to remove them from the ignore list:",
                .Dock = DockStyle.Top,
                .Height = 40,
                .Padding = New Padding(10)
            }
            form.Controls.Add(lblInfo)

            Dim clb As New CheckedListBox() With {
                .Dock = DockStyle.Fill,
                .CheckOnClick = True
            }
            form.Controls.Add(clb)

            ' Load all ignored items
            Dim allIgnored As New List(Of Tuple(Of String, String))() ' (file, key)

            For Each fileName In {"redink.ini", "allmodels.ini", "specialservices.ini"}
                Dim ignoreList = GetIgnoreListForFile(fileName)
                For Each key In ignoreList
                    allIgnored.Add(Tuple.Create(fileName, key))
                    clb.Items.Add($"[{fileName}] {key}", True)
                Next
            Next

            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .Height = 50,
                .Padding = New Padding(10)
            }
            form.Controls.Add(pnlButtons)

            Dim btnCancel As New Button() With {.Text = "Cancel", .AutoSize = True}
            Dim btnSave As New Button() With {.Text = "Save Changes", .AutoSize = True}

            pnlButtons.Controls.Add(btnCancel)
            pnlButtons.Controls.Add(btnSave)

            AddHandler btnCancel.Click, Sub() form.Close()

            AddHandler btnSave.Click, Sub()
                                          ' Rebuild ignore lists based on checked items
                                          Dim newIgnoreLists As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase) From {
                                              {"redink.ini", New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)},
                                              {"allmodels.ini", New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)},
                                              {"specialservices.ini", New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)}
                                          }

                                          For i As Integer = 0 To clb.Items.Count - 1
                                              If clb.GetItemChecked(i) Then
                                                  Dim item = allIgnored(i)
                                                  newIgnoreLists(item.Item1).Add(item.Item2)
                                              End If
                                          Next

                                          For Each kvp In newIgnoreLists
                                              SaveIgnoreListForFile(kvp.Key, kvp.Value)
                                          Next

                                          ShowCustomMessageBox("Ignore list updated successfully.")
                                          form.Close()
                                      End Sub

            clb.BringToFront()
            form.ShowDialog()
        End Sub

#End Region

#Region "Update Approval Dialog"

        ''' <summary>
        ''' Shows the update approval dialog with all detected changes.
        ''' </summary>
        Private Shared Function ShowUpdateApprovalDialog(changes As List(Of IniParameterChange)) As UpdateApprovalResult
            Dim result As UpdateApprovalResult = UpdateApprovalResult.Cancel

            Dim form As New Form() With {
                .Text = $"{AN} - Configuration Updates Available",
                .Size = New Size(900, 600),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(700, 400)
            }

            Try
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Header
            Dim lblHeader As New Label() With {
                .Text = "The following configuration updates are available. Review and select which changes to apply:" & vbCrLf &
                        "(Items shown in red contain URL or path changes and are not selected by default for security reasons)",
                .Dock = DockStyle.Top,
                .Height = 50,
                .Padding = New Padding(10)
            }
            form.Controls.Add(lblHeader)

            ' DataGridView for changes
            Dim dgv As New DataGridView() With {
                .Dock = DockStyle.Fill,
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .MultiSelect = False,
                .RowHeadersVisible = False
            }

            ' Columns
            Dim colApply As New DataGridViewCheckBoxColumn() With {
                .HeaderText = "Apply",
                .Name = "colApply",
                .Width = 50,
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            }
            Dim colFile As New DataGridViewTextBoxColumn() With {
                .HeaderText = "File",
                .Name = "colFile",
                .ReadOnly = True,
                .Width = 100,
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            }
            Dim colSegment As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Segment",
                .Name = "colSegment",
                .ReadOnly = True,
                .Width = 100,
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            }
            Dim colKey As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Parameter",
                .Name = "colKey",
                .ReadOnly = True,
                .Width = 120,
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            }
            Dim colOld As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Current Value",
                .Name = "colOld",
                .ReadOnly = True
            }
            Dim colNew As New DataGridViewTextBoxColumn() With {
                .HeaderText = "New Value",
                .Name = "colNew",
                .ReadOnly = True
            }

            dgv.Columns.AddRange(colApply, colFile, colSegment, colKey, colOld, colNew)

            ' Add rows
            For Each change In changes
                Dim rowIndex = dgv.Rows.Add(
                    change.IsSelected,
                    change.IniFile,
                    If(change.SegmentName, ""),
                    change.ParameterKey,
                    TruncateValue(change.OldValue, 100),
                    TruncateValue(change.NewValue, 100)
                )

                ' Color suspicious rows red
                If change.IsSuspicious Then
                    dgv.Rows(rowIndex).DefaultCellStyle.ForeColor = Color.Red
                End If
            Next

            ' Handle checkbox changes
            AddHandler dgv.CellValueChanged, Sub(sender, e)
                                                 If e.ColumnIndex = 0 AndAlso e.RowIndex >= 0 Then
                                                     changes(e.RowIndex).IsSelected = CBool(dgv.Rows(e.RowIndex).Cells(0).Value)
                                                 End If
                                             End Sub

            AddHandler dgv.CurrentCellDirtyStateChanged, Sub()
                                                             If dgv.IsCurrentCellDirty Then
                                                                 dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                                                             End If
                                                         End Sub

            form.Controls.Add(dgv)

            ' Buttons panel
            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .Height = 50,
                .Padding = New Padding(10)
            }
            form.Controls.Add(pnlButtons)

            Dim btnReject As New Button() With {.Text = "Reject All", .AutoSize = True}
            Dim btnApprove As New Button() With {.Text = "Approve Selected", .AutoSize = True}

            pnlButtons.Controls.Add(btnReject)
            pnlButtons.Controls.Add(btnApprove)

            AddHandler btnApprove.Click, Sub()
                                             result = UpdateApprovalResult.Approve
                                             form.Close()
                                         End Sub

            AddHandler btnReject.Click, Sub()
                                            ' Deselect all
                                            For Each change In changes
                                                change.IsSelected = False
                                            Next
                                            result = UpdateApprovalResult.Reject
                                            form.Close()
                                        End Sub

            dgv.BringToFront()
            form.ShowDialog()

            Return result
        End Function

        ''' <summary>
        ''' Shows confirmation dialog for adding rejected items to ignore list.
        ''' </summary>
        Private Shared Sub ShowIgnoreConfirmationDialog(rejectedChanges As List(Of IniParameterChange))
            If rejectedChanges Is Nothing OrElse rejectedChanges.Count = 0 Then Return

            Dim form As New Form() With {
                .Text = $"{AN} - Ignore Future Updates?",
                .Size = New Size(700, 450),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F)
            }

            Try
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Dim lblInfo As New Label() With {
                .Text = "The following parameters were not approved. Select which ones to ignore in future update checks:",
                .Dock = DockStyle.Top,
                .Height = 40,
                .Padding = New Padding(10)
            }
            form.Controls.Add(lblInfo)

            Dim clb As New CheckedListBox() With {
                .Dock = DockStyle.Fill,
                .CheckOnClick = True
            }
            form.Controls.Add(clb)

            For Each change In rejectedChanges
                clb.Items.Add($"[{change.IniFile}]{If(String.IsNullOrEmpty(change.SegmentName), "", $"[{change.SegmentName}]")}.{change.ParameterKey}", False)
            Next

            Dim pnlButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .Height = 50,
                .Padding = New Padding(10)
            }
            form.Controls.Add(pnlButtons)

            Dim btnAbort As New Button() With {.Text = "Don't Ignore Any", .AutoSize = True}
            Dim btnIgnore As New Button() With {.Text = "Ignore Selected", .AutoSize = True}

            pnlButtons.Controls.Add(btnAbort)
            pnlButtons.Controls.Add(btnIgnore)

            AddHandler btnAbort.Click, Sub() form.Close()

            AddHandler btnIgnore.Click, Sub()
                                            Dim toIgnore As New List(Of IniParameterChange)()
                                            For i As Integer = 0 To clb.Items.Count - 1
                                                If clb.GetItemChecked(i) Then
                                                    toIgnore.Add(rejectedChanges(i))
                                                End If
                                            Next

                                            If toIgnore.Count > 0 Then
                                                AddToIgnoreList(toIgnore)
                                                ShowCustomMessageBox($"{toIgnore.Count} parameter(s) will be ignored in future update checks.")
                                            End If

                                            form.Close()
                                        End Sub

            clb.BringToFront()
            form.ShowDialog()
        End Sub

        ''' <summary>
        ''' Truncates a value for display purposes.
        ''' </summary>
        Private Shared Function TruncateValue(value As String, maxLen As Integer) As String
            If String.IsNullOrEmpty(value) Then Return ""
            If value.Length <= maxLen Then Return value
            Return value.Substring(0, maxLen - 3) & "..."
        End Function

#End Region

#Region "Apply Updates"

        ''' <summary>
        ''' Applies approved changes to the INI files.
        ''' </summary>
        Private Shared Sub ApplyApprovedUpdates(changes As List(Of IniParameterChange), context As ISharedContext)
            ' Group changes by file
            Dim byFile = changes.GroupBy(Function(c) c.IniFile)

            For Each fileGroup In byFile
                Dim filePath As String = Nothing
                Dim fileName = fileGroup.Key

                ' Determine the actual file path based on the filename
                ' Check if it's the main config or one of the user-defined paths
                Dim mainIniPath = GetDefaultINIPath(context.RDV)
                Dim mainIniName = Path.GetFileName(mainIniPath)

                If fileName.Equals(mainIniName, StringComparison.OrdinalIgnoreCase) Then
                    filePath = mainIniPath
                ElseIf Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                    Dim altPath = ExpandEnvironmentVariables(context.INI_AlternateModelPath)
                    If Path.GetFileName(altPath).Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                        filePath = altPath
                    End If
                End If

                If filePath Is Nothing AndAlso Not String.IsNullOrWhiteSpace(context.INI_SpecialServicePath) Then
                    Dim svcPath = ExpandEnvironmentVariables(context.INI_SpecialServicePath)
                    If Path.GetFileName(svcPath).Equals(fileName, StringComparison.OrdinalIgnoreCase) Then
                        filePath = svcPath
                    End If
                End If

                If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then Continue For

                Try
                    ' Create backup
                    RenameFileToBak(filePath)

                    ' Read current content from backup
                    Dim lines = File.ReadAllLines(filePath & ".bak", Encoding.UTF8).ToList()
                    Dim updatedLines As New List(Of String)()

                    ' Build lookup of changes by segment
                    Dim changesBySegment = fileGroup.GroupBy(Function(c) If(c.SegmentName, "")).
                        ToDictionary(Function(g) g.Key,
                                     Function(g) g.ToDictionary(Function(c) c.ParameterKey, StringComparer.OrdinalIgnoreCase))

                    Dim currentSegment As String = ""
                    Dim usedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                    For Each line In lines
                        Dim trimmed = line.Trim()

                        ' Track segment changes
                        If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") Then
                            currentSegment = trimmed.Substring(1, trimmed.Length - 2).Trim()
                            usedKeys.Clear()
                            updatedLines.Add(line)
                            Continue For
                        End If

                        ' Check if this is a key=value line
                        Dim eqIndex = trimmed.IndexOf("="c)
                        If eqIndex > 0 AndAlso Not trimmed.StartsWith(";") Then
                            Dim key = trimmed.Substring(0, eqIndex).Trim()

                            ' Check if this key needs updating
                            If changesBySegment.ContainsKey(currentSegment) AndAlso
                               changesBySegment(currentSegment).ContainsKey(key) Then
                                Dim change = changesBySegment(currentSegment)(key)
                                updatedLines.Add($"{key} = {change.NewValue}")
                                usedKeys.Add(key)
                                Continue For
                            End If
                        End If

                        updatedLines.Add(line)
                    Next

                    ' Write updated content
                    File.WriteAllLines(filePath, updatedLines, Encoding.UTF8)

                Catch ex As Exception
                    Debug.WriteLine($"Error applying updates to {filePath}: {ex.Message}")
                    ShowCustomMessageBox($"Failed to update {fileName}: {ex.Message}")
                    ' Try to restore backup
                    Try
                        If File.Exists(filePath & ".bak") Then
                            If File.Exists(filePath) Then File.Delete(filePath)
                            File.Move(filePath & ".bak", filePath)
                        End If
                    Catch
                    End Try
                End Try
            Next
        End Sub

#End Region

#Region "Ed25519 Digital Signature"

        ''' <summary>
        ''' Generates a new Ed25519 keypair for signing update files.
        ''' </summary>
        ''' <returns>Tuple of (Base64 Public Key, Base64 Private Key)</returns>
        Public Shared Function GenerateEd25519KeyPair() As (PublicKey As String, PrivateKey As String)
            Try
                Dim keyGen As New Ed25519KeyPairGenerator()
                keyGen.Init(New Ed25519KeyGenerationParameters(New SecureRandom()))
                Dim keyPair = keyGen.GenerateKeyPair()

                Dim publicKey = DirectCast(keyPair.Public, Ed25519PublicKeyParameters)
                Dim privateKey = DirectCast(keyPair.Private, Ed25519PrivateKeyParameters)

                Dim pubBytes = publicKey.GetEncoded()
                Dim privBytes = privateKey.GetEncoded()

                Return (System.Convert.ToBase64String(pubBytes), System.Convert.ToBase64String(privBytes))

            Catch ex As Exception
                Debug.WriteLine($"Error generating Ed25519 keypair: {ex.Message}")
                Return (Nothing, Nothing)
            End Try
        End Function

        ''' <summary>
        ''' Signs a file using Ed25519 and creates a .sig file alongside it.
        ''' </summary>
        ''' <param name="filePath">Path to the file to sign.</param>
        ''' <param name="base64PrivateKey">Base64-encoded Ed25519 private key.</param>
        ''' <returns>True if signing succeeded.</returns>
        Public Shared Function SignUpdateFile(filePath As String, base64PrivateKey As String) As Boolean
            Try
                If Not File.Exists(filePath) Then
                    ShowCustomMessageBox($"File not found: {filePath}")
                    Return False
                End If

                Dim privKeyBytes = System.Convert.FromBase64String(base64PrivateKey)
                Dim privateKey As New Ed25519PrivateKeyParameters(privKeyBytes, 0)

                Dim content = File.ReadAllBytes(filePath)

                Dim signer As New Ed25519Signer()
                signer.Init(True, privateKey)
                signer.BlockUpdate(content, 0, content.Length)
                Dim signature = signer.GenerateSignature()

                Dim sigPath = filePath & ".sig"
                File.WriteAllText(sigPath, System.Convert.ToBase64String(signature))

                Return True

            Catch ex As Exception
                ShowCustomMessageBox($"Error signing file: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Verifies an Ed25519 signature.
        ''' </summary>
        Private Shared Function VerifyEd25519Signature(content As String, signatureBase64 As String, publicKeyBase64 As String) As Boolean
            Try
                Dim pubKeyBytes = System.Convert.FromBase64String(publicKeyBase64)
                Dim publicKey As New Ed25519PublicKeyParameters(pubKeyBytes, 0)

                Dim signatureBytes = System.Convert.FromBase64String(signatureBase64)
                Dim contentBytes = Encoding.UTF8.GetBytes(content)

                Dim verifier As New Ed25519Signer()
                verifier.Init(False, publicKey)
                verifier.BlockUpdate(contentBytes, 0, contentBytes.Length)

                Return verifier.VerifySignature(signatureBytes)

            Catch ex As Exception
                Debug.WriteLine($"Signature verification error: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Verifies a signature file (.sig) against its source file.
        ''' </summary>
        Public Shared Function VerifySignatureFile(filePath As String, publicKeyBase64 As String) As Boolean
            Try
                If Not File.Exists(filePath) Then
                    ShowCustomMessageBox($"File not found: {filePath}")
                    Return False
                End If

                Dim sigPath = filePath & ".sig"
                If Not File.Exists(sigPath) Then
                    ShowCustomMessageBox($"Signature file not found: {sigPath}")
                    Return False
                End If

                Dim content = File.ReadAllText(filePath, Encoding.UTF8)
                Dim signature = File.ReadAllText(sigPath, Encoding.UTF8).Trim()

                Return VerifyEd25519Signature(content, signature, publicKeyBase64)

            Catch ex As Exception
                ShowCustomMessageBox($"Error verifying signature: {ex.Message}")
                Return False
            End Try
        End Function

#End Region


#Region "Signature Management UI"

        ''' <summary>
        ''' Shows the signature management UI for generating keys, signing files, and verifying signatures.
        ''' </summary>
        Public Shared Sub ShowSignatureManagementDialog()
            Dim form As New Form() With {
                .Text = $"{AN} - Update Signature Management",
                .Size = New Size(750, 550),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F),
                .MinimumSize = New Size(600, 450)
            }

            Try
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Dim tabControl As New TabControl() With {.Dock = DockStyle.Fill}
            form.Controls.Add(tabControl)

            ' =====================================================================
            ' Tab 1: Generate Keypair
            ' =====================================================================
            Dim tabGenerate As New TabPage("Generate Keypair")
            tabControl.TabPages.Add(tabGenerate)

            Dim pnlGenerate As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 2,
                .RowCount = 6
            }
            pnlGenerate.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 200))
            pnlGenerate.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            tabGenerate.Controls.Add(pnlGenerate)

            ' Row 0: Info label
            Dim lblGenInfo As New Label() With {
                .Text = "Generate a new Ed25519 keypair for signing update files." & vbCrLf &
                        "The public key goes into UpdateSource. Keep the private key secure!",
                .Dock = DockStyle.Fill,
                .AutoSize = False,
                .Height = 50
            }
            pnlGenerate.Controls.Add(lblGenInfo, 0, 0)
            pnlGenerate.SetColumnSpan(lblGenInfo, 2)

            ' Row 1: Generate button
            Dim btnGenerate As New Button() With {
                .Text = "Generate New Keypair",
                .Dock = DockStyle.Left,
                .Width = 180,
                .Height = 30
            }
            pnlGenerate.Controls.Add(btnGenerate, 0, 1)

            ' Row 2: Public Key label and textbox
            Dim lblPubKey As New Label() With {
                .Text = "Public Key (for UpdateSource):",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlGenerate.Controls.Add(lblPubKey, 0, 2)

            Dim txtPubKey As New TextBox() With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .BackColor = SystemColors.Window
            }
            pnlGenerate.Controls.Add(txtPubKey, 1, 2)

            ' Row 3: Private Key label and textbox
            Dim lblPrivKey As New Label() With {
                .Text = "Private Key (KEEP SECRET!):",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft,
                .ForeColor = Color.DarkRed
            }
            pnlGenerate.Controls.Add(lblPrivKey, 0, 3)

            Dim txtPrivKey As New TextBox() With {
                .Dock = DockStyle.Fill,
                .ReadOnly = True,
                .BackColor = SystemColors.Window,
                .UseSystemPasswordChar = False
            }
            pnlGenerate.Controls.Add(txtPrivKey, 1, 3)

            ' Row 4: Copy buttons
            Dim pnlCopyButtons As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .FlowDirection = FlowDirection.LeftToRight
            }
            pnlGenerate.Controls.Add(pnlCopyButtons, 1, 4)

            Dim btnCopyPub As New Button() With {.Text = "Copy Public Key", .AutoSize = True}
            Dim btnCopyPriv As New Button() With {.Text = "Copy Private Key", .AutoSize = True}
            Dim btnCopyBoth As New Button() With {.Text = "Copy Both to Clipboard", .AutoSize = True}
            pnlCopyButtons.Controls.AddRange({btnCopyPub, btnCopyPriv, btnCopyBoth})

            AddHandler btnGenerate.Click, Sub()
                                              Dim keys = GenerateEd25519KeyPair()
                                              If keys.PublicKey IsNot Nothing Then
                                                  txtPubKey.Text = keys.PublicKey
                                                  txtPrivKey.Text = keys.PrivateKey
                                                  ShowCustomMessageBox("Keypair generated successfully!" & vbCrLf & vbCrLf &
                                                      "Public Key: Use this in the UpdateSource parameter of your INI segments." & vbCrLf &
                                                      "Private Key: Store securely and use for signing update files.")
                                              End If
                                          End Sub

            AddHandler btnCopyPub.Click, Sub()
                                             If Not String.IsNullOrEmpty(txtPubKey.Text) Then
                                                 Try
                                                     Clipboard.SetText(txtPubKey.Text)
                                                     ShowCustomMessageBox("Public key copied to clipboard.")
                                                 Catch ex As Exception
                                                     ShowCustomMessageBox("Failed to copy: " & ex.Message)
                                                 End Try
                                             End If
                                         End Sub

            AddHandler btnCopyPriv.Click, Sub()
                                              If Not String.IsNullOrEmpty(txtPrivKey.Text) Then
                                                  Try
                                                      Clipboard.SetText(txtPrivKey.Text)
                                                      ShowCustomMessageBox("Private key copied to clipboard. Keep it secure!")
                                                  Catch ex As Exception
                                                      ShowCustomMessageBox("Failed to copy: " & ex.Message)
                                                  End Try
                                              End If
                                          End Sub

            AddHandler btnCopyBoth.Click, Sub()
                                              If Not String.IsNullOrEmpty(txtPubKey.Text) Then
                                                  Try
                                                      Dim text = $"=== Ed25519 KEYPAIR ==={vbCrLf}{vbCrLf}" &
                                                                 $"PUBLIC KEY (for UpdateSource):{vbCrLf}{txtPubKey.Text}{vbCrLf}{vbCrLf}" &
                                                                 $"PRIVATE KEY (KEEP SECRET - for signing):{vbCrLf}{txtPrivKey.Text}{vbCrLf}"
                                                      Clipboard.SetText(text)
                                                      ShowCustomMessageBox("Both keys copied to clipboard.")
                                                  Catch ex As Exception
                                                      ShowCustomMessageBox("Failed to copy: " & ex.Message)
                                                  End Try
                                              End If
                                          End Sub

            ' =====================================================================
            ' Tab 2: Sign File
            ' =====================================================================
            Dim tabSign As New TabPage("Sign File")
            tabControl.TabPages.Add(tabSign)

            Dim pnlSign As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 3,
                .RowCount = 6
            }
            pnlSign.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            pnlSign.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            pnlSign.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 40))
            tabSign.Controls.Add(pnlSign)

            ' Row 0: Info
            Dim lblSignInfo As New Label() With {
                .Text = "Sign an update INI file to create a .sig signature file alongside it." & vbCrLf &
                        "The signature file must be placed next to the update source file.",
                .Dock = DockStyle.Fill,
                .Height = 50
            }
            pnlSign.Controls.Add(lblSignInfo, 0, 0)
            pnlSign.SetColumnSpan(lblSignInfo, 3)

            ' Row 1: File to sign
            Dim lblSignFile As New Label() With {
                .Text = "File to Sign:",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlSign.Controls.Add(lblSignFile, 0, 1)

            Dim txtSignFile As New TextBox() With {.Dock = DockStyle.Fill}
            pnlSign.Controls.Add(txtSignFile, 1, 1)

            Dim btnBrowseSign As New Button() With {.Text = "...", .Dock = DockStyle.Fill}
            pnlSign.Controls.Add(btnBrowseSign, 2, 1)

            ' Row 2: Private key
            Dim lblSignPrivKey As New Label() With {
                .Text = "Private Key:",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlSign.Controls.Add(lblSignPrivKey, 0, 2)

            Dim txtSignPrivKey As New TextBox() With {
                .Dock = DockStyle.Fill,
                .UseSystemPasswordChar = True
            }
            pnlSign.Controls.Add(txtSignPrivKey, 1, 2)
            pnlSign.SetColumnSpan(txtSignPrivKey, 2)

            ' Row 3: Show/hide password toggle
            Dim chkShowPrivKey As New CheckBox() With {
                .Text = "Show private key",
                .Dock = DockStyle.Left,
                .AutoSize = True
            }
            pnlSign.Controls.Add(chkShowPrivKey, 1, 3)

            AddHandler chkShowPrivKey.CheckedChanged, Sub()
                                                          txtSignPrivKey.UseSystemPasswordChar = Not chkShowPrivKey.Checked
                                                      End Sub

            ' Row 4: Sign button
            Dim btnSign As New Button() With {
                .Text = "Sign File",
                .Width = 120,
                .Height = 35,
                .Dock = DockStyle.Left
            }
            pnlSign.Controls.Add(btnSign, 1, 4)

            ' Row 5: Result
            Dim lblSignResult As New Label() With {
                .Text = "",
                .Dock = DockStyle.Fill,
                .ForeColor = Color.DarkGreen
            }
            pnlSign.Controls.Add(lblSignResult, 0, 5)
            pnlSign.SetColumnSpan(lblSignResult, 3)

            AddHandler btnBrowseSign.Click, Sub()
                                                Using ofd As New OpenFileDialog()
                                                    ofd.Title = "Select File to Sign"
                                                    ofd.Filter = "INI Files|*.ini|Text Files|*.txt|All Files|*.*"
                                                    If ofd.ShowDialog() = DialogResult.OK Then
                                                        txtSignFile.Text = ofd.FileName
                                                    End If
                                                End Using
                                            End Sub

            AddHandler btnSign.Click, Sub()
                                          lblSignResult.Text = ""
                                          lblSignResult.ForeColor = Color.DarkGreen

                                          If String.IsNullOrWhiteSpace(txtSignFile.Text) Then
                                              ShowCustomMessageBox("Please select a file to sign.")
                                              Return
                                          End If

                                          If Not File.Exists(txtSignFile.Text) Then
                                              ShowCustomMessageBox("The selected file does not exist.")
                                              Return
                                          End If

                                          If String.IsNullOrWhiteSpace(txtSignPrivKey.Text) Then
                                              ShowCustomMessageBox("Please enter the private key.")
                                              Return
                                          End If

                                          Try
                                              If SignUpdateFile(txtSignFile.Text, txtSignPrivKey.Text.Trim()) Then
                                                  lblSignResult.Text = $"✓ Signature created: {txtSignFile.Text}.sig"
                                                  lblSignResult.ForeColor = Color.DarkGreen
                                                  ShowCustomMessageBox($"File signed successfully!{vbCrLf}{vbCrLf}" &
                                                      $"Signature saved to:{vbCrLf}{txtSignFile.Text}.sig{vbCrLf}{vbCrLf}" &
                                                      "Upload both the INI file and its .sig file to your update location.")
                                              Else
                                                  lblSignResult.Text = "✗ Signing failed"
                                                  lblSignResult.ForeColor = Color.DarkRed
                                              End If
                                          Catch ex As Exception
                                              lblSignResult.Text = $"✗ Error: {ex.Message}"
                                              lblSignResult.ForeColor = Color.DarkRed
                                          End Try
                                      End Sub

            ' =====================================================================
            ' Tab 3: Verify Signature
            ' =====================================================================
            Dim tabVerify As New TabPage("Verify Signature")
            tabControl.TabPages.Add(tabVerify)

            Dim pnlVerify As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15),
                .ColumnCount = 3,
                .RowCount = 6
            }
            pnlVerify.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
            pnlVerify.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            pnlVerify.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 40))
            tabVerify.Controls.Add(pnlVerify)

            ' Row 0: Info
            Dim lblVerifyInfo As New Label() With {
                .Text = "Verify that a file's signature is valid using the public key." & vbCrLf &
                        "The .sig file must exist alongside the file being verified.",
                .Dock = DockStyle.Fill,
                .Height = 50
            }
            pnlVerify.Controls.Add(lblVerifyInfo, 0, 0)
            pnlVerify.SetColumnSpan(lblVerifyInfo, 3)

            ' Row 1: File to verify
            Dim lblVerifyFile As New Label() With {
                .Text = "File to Verify:",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlVerify.Controls.Add(lblVerifyFile, 0, 1)

            Dim txtVerifyFile As New TextBox() With {.Dock = DockStyle.Fill}
            pnlVerify.Controls.Add(txtVerifyFile, 1, 1)

            Dim btnBrowseVerify As New Button() With {.Text = "...", .Dock = DockStyle.Fill}
            pnlVerify.Controls.Add(btnBrowseVerify, 2, 1)

            ' Row 2: Public key
            Dim lblVerifyPubKey As New Label() With {
                .Text = "Public Key:",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleLeft
            }
            pnlVerify.Controls.Add(lblVerifyPubKey, 0, 2)

            Dim txtVerifyPubKey As New TextBox() With {.Dock = DockStyle.Fill}
            pnlVerify.Controls.Add(txtVerifyPubKey, 1, 2)
            pnlVerify.SetColumnSpan(txtVerifyPubKey, 2)

            ' Row 3: Verify button
            Dim btnVerify As New Button() With {
                .Text = "Verify Signature",
                .Width = 120,
                .Height = 35,
                .Dock = DockStyle.Left
            }
            pnlVerify.Controls.Add(btnVerify, 1, 3)

            ' Row 4: Result
            Dim lblVerifyResult As New Label() With {
                .Text = "",
                .Dock = DockStyle.Fill,
                .Font = New Font("Segoe UI", 11.0F, FontStyle.Bold)
            }
            pnlVerify.Controls.Add(lblVerifyResult, 0, 4)
            pnlVerify.SetColumnSpan(lblVerifyResult, 3)

            AddHandler btnBrowseVerify.Click, Sub()
                                                  Using ofd As New OpenFileDialog()
                                                      ofd.Title = "Select File to Verify"
                                                      ofd.Filter = "INI Files|*.ini|Text Files|*.txt|All Files|*.*"
                                                      If ofd.ShowDialog() = DialogResult.OK Then
                                                          txtVerifyFile.Text = ofd.FileName
                                                      End If
                                                  End Using
                                              End Sub

            AddHandler btnVerify.Click, Sub()
                                            lblVerifyResult.Text = ""

                                            If String.IsNullOrWhiteSpace(txtVerifyFile.Text) Then
                                                ShowCustomMessageBox("Please select a file to verify.")
                                                Return
                                            End If

                                            If Not File.Exists(txtVerifyFile.Text) Then
                                                ShowCustomMessageBox("The selected file does not exist.")
                                                Return
                                            End If

                                            Dim sigPath = txtVerifyFile.Text & ".sig"
                                            If Not File.Exists(sigPath) Then
                                                ShowCustomMessageBox($"Signature file not found:{vbCrLf}{sigPath}")
                                                Return
                                            End If

                                            If String.IsNullOrWhiteSpace(txtVerifyPubKey.Text) Then
                                                ShowCustomMessageBox("Please enter the public key.")
                                                Return
                                            End If

                                            Try
                                                Dim isValid = VerifySignatureFile(txtVerifyFile.Text, txtVerifyPubKey.Text.Trim())
                                                If isValid Then
                                                    lblVerifyResult.Text = "✓ SIGNATURE VALID - File is authentic"
                                                    lblVerifyResult.ForeColor = Color.DarkGreen
                                                Else
                                                    lblVerifyResult.Text = "✗ SIGNATURE INVALID - File may have been modified!"
                                                    lblVerifyResult.ForeColor = Color.DarkRed
                                                End If
                                            Catch ex As Exception
                                                lblVerifyResult.Text = $"✗ Verification failed: {ex.Message}"
                                                lblVerifyResult.ForeColor = Color.DarkRed
                                            End Try
                                        End Sub

            ' =====================================================================
            ' Tab 4: Help / Instructions
            ' =====================================================================
            Dim tabHelp As New TabPage("Help")
            tabControl.TabPages.Add(tabHelp)

            Dim txtHelp As New TextBox() With {
                .Dock = DockStyle.Fill,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Vertical,
                .BackColor = SystemColors.Window,
                .Font = New Font("Segoe UI", 9.5F),
                .Text = GetSignatureHelpText()
            }
            tabHelp.Controls.Add(txtHelp)

            form.ShowDialog()
        End Sub

        ''' <summary>
        ''' Returns help text for the signature management dialog.
        ''' </summary>
        Private Shared Function GetSignatureHelpText() As String
            Return $"=== {AN} Update Signature System ===" & vbCrLf & vbCrLf &
                "This tool uses Ed25519 digital signatures to ensure the authenticity " &
                "and integrity of configuration updates. Ed25519 is a modern, secure " &
                "signature algorithm that provides strong protection against tampering." & vbCrLf & vbCrLf &
                "=== HOW IT WORKS ===" & vbCrLf & vbCrLf &
                "1. GENERATE KEYPAIR" & vbCrLf &
                "   - Go to 'Generate Keypair' tab and click 'Generate New Keypair'" & vbCrLf &
                "   - You'll receive a Public Key and a Private Key" & vbCrLf &
                "   - Store the Private Key securely (e.g., in a password manager)" & vbCrLf &
                "   - The Public Key will be included in UpdateSource entries" & vbCrLf & vbCrLf &
                "2. CONFIGURE UPDATE SOURCE" & vbCrLf &
                "   In your INI files, configure UpdateSource as:" & vbCrLf &
                "   UpdateSource = path; keys; public_key" & vbCrLf & vbCrLf &
                "   Example:" & vbCrLf &
                "   UpdateSource = https://example.com/updates/models.ini; all; MCow..." & vbCrLf & vbCrLf &
                "   Where:" & vbCrLf &
                "   - path: URL or file path to the update INI file" & vbCrLf &
                "   - keys: 'all' or comma-separated list of keys to update" & vbCrLf &
                "   - public_key: Base64-encoded Ed25519 public key" & vbCrLf & vbCrLf &
                "3. SIGN UPDATE FILES" & vbCrLf &
                "   - Create/modify your update INI file" & vbCrLf &
                "   - Go to 'Sign File' tab" & vbCrLf &
                "   - Select the file and enter your Private Key" & vbCrLf &
                "   - Click 'Sign File' to create the .sig file" & vbCrLf &
                "   - Upload BOTH the INI file and its .sig file to your server" & vbCrLf & vbCrLf &
                "4. VERIFY SIGNATURES (Optional)" & vbCrLf &
                "   - Use the 'Verify Signature' tab to manually test" & vbCrLf &
                "   - This is useful for troubleshooting" & vbCrLf & vbCrLf &
                "=== SECURITY NOTES ===" & vbCrLf & vbCrLf &
                "• NEVER share your Private Key" & vbCrLf &
                "• The Private Key is needed only for signing (administrator)" & vbCrLf &
                "• The Public Key is safe to distribute in UpdateSource" & vbCrLf &
                "• If your Private Key is compromised, generate a new keypair" & vbCrLf &
                "  and update all UpdateSource entries with the new Public Key" & vbCrLf & vbCrLf &
                "=== FILE STRUCTURE ===" & vbCrLf & vbCrLf &
                "When you sign 'updates.ini', the system creates 'updates.ini.sig'" & vbCrLf &
                "Both files must be uploaded to the same location:" & vbCrLf & vbCrLf &
                "   https://example.com/updates/models.ini" & vbCrLf &
                "   https://example.com/updates/models.ini.sig" & vbCrLf & vbCrLf &
                "The update checker downloads both files and verifies the signature " &
                "before applying any changes."
        End Function

#End Region

#Region "Batch Signing Utility"

        ''' <summary>
        ''' Signs multiple files at once using the same private key.
        ''' </summary>
        ''' <param name="filePaths">Array of file paths to sign.</param>
        ''' <param name="base64PrivateKey">Base64-encoded Ed25519 private key.</param>
        ''' <returns>Dictionary of results (file path -> success/error message).</returns>
        Public Shared Function BatchSignFiles(filePaths As String(), base64PrivateKey As String) As Dictionary(Of String, String)
            Dim results As New Dictionary(Of String, String)()

            If filePaths Is Nothing OrElse filePaths.Length = 0 Then
                Return results
            End If

            For Each filePath In filePaths
                Try
                    If SignUpdateFile(filePath, base64PrivateKey) Then
                        results(filePath) = "Success"
                    Else
                        results(filePath) = "Failed (unknown error)"
                    End If
                Catch ex As Exception
                    results(filePath) = $"Error: {ex.Message}"
                End Try
            Next

            Return results
        End Function

        ''' <summary>
        ''' Shows a batch signing dialog for signing multiple files.
        ''' </summary>
        Public Shared Sub ShowBatchSigningDialog()
            Dim form As New Form() With {
                .Text = $"{AN} - Batch Sign Files",
                .Size = New Size(700, 500),
                .StartPosition = FormStartPosition.CenterScreen,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .Font = New Font("Segoe UI", 9.0F)
            }

            Try
                Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
                form.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            Dim pnlTop As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 80,
                .Padding = New Padding(10)
            }
            form.Controls.Add(pnlTop)

            Dim lblPrivKey As New Label() With {
                .Text = "Private Key:",
                .Location = New Point(10, 10),
                .AutoSize = True
            }
            pnlTop.Controls.Add(lblPrivKey)

            Dim txtPrivKey As New TextBox() With {
                .Location = New Point(100, 7),
                .Width = 450,
                .UseSystemPasswordChar = True
            }
            pnlTop.Controls.Add(txtPrivKey)

            Dim chkShowKey As New CheckBox() With {
                .Text = "Show",
                .Location = New Point(560, 9),
                .AutoSize = True
            }
            pnlTop.Controls.Add(chkShowKey)

            AddHandler chkShowKey.CheckedChanged, Sub()
                                                      txtPrivKey.UseSystemPasswordChar = Not chkShowKey.Checked
                                                  End Sub

            Dim btnAddFiles As New Button() With {
                .Text = "Add Files...",
                .Location = New Point(10, 40),
                .AutoSize = True
            }
            pnlTop.Controls.Add(btnAddFiles)

            Dim btnClear As New Button() With {
                .Text = "Clear List",
                .Location = New Point(100, 40),
                .AutoSize = True
            }
            pnlTop.Controls.Add(btnClear)

            Dim lbFiles As New ListBox() With {
                .Dock = DockStyle.Fill,
                .SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            }
            form.Controls.Add(lbFiles)

            Dim pnlBottom As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .Height = 50,
                .Padding = New Padding(10)
            }
            form.Controls.Add(pnlBottom)

            Dim btnClose As New Button() With {.Text = "Close", .AutoSize = True}
            Dim btnSignAll As New Button() With {.Text = "Sign All Files", .AutoSize = True}

            pnlBottom.Controls.Add(btnClose)
            pnlBottom.Controls.Add(btnSignAll)

            AddHandler btnAddFiles.Click, Sub()
                                              Using ofd As New OpenFileDialog()
                                                  ofd.Title = "Select Files to Sign"
                                                  ofd.Filter = "INI Files|*.ini|Text Files|*.txt|All Files|*.*"
                                                  ofd.Multiselect = True
                                                  If ofd.ShowDialog() = DialogResult.OK Then
                                                      For Each f In ofd.FileNames
                                                          If Not lbFiles.Items.Contains(f) Then
                                                              lbFiles.Items.Add(f)
                                                          End If
                                                      Next
                                                  End If
                                              End Using
                                          End Sub

            AddHandler btnClear.Click, Sub() lbFiles.Items.Clear()

            AddHandler btnClose.Click, Sub() form.Close()

            AddHandler btnSignAll.Click, Sub()
                                             If lbFiles.Items.Count = 0 Then
                                                 ShowCustomMessageBox("No files to sign. Add files first.")
                                                 Return
                                             End If

                                             If String.IsNullOrWhiteSpace(txtPrivKey.Text) Then
                                                 ShowCustomMessageBox("Please enter the private key.")
                                                 Return
                                             End If

                                             Dim files = lbFiles.Items.Cast(Of String)().ToArray()
                                             Dim results = BatchSignFiles(files, txtPrivKey.Text.Trim())

                                             Dim sb As New StringBuilder()
                                             sb.AppendLine("Batch Signing Results:")
                                             sb.AppendLine()

                                             Dim successCount = 0
                                             For Each kvp In results
                                                 Dim fileName = Path.GetFileName(kvp.Key)
                                                 If kvp.Value = "Success" Then
                                                     sb.AppendLine($"✓ {fileName}")
                                                     successCount += 1
                                                 Else
                                                     sb.AppendLine($"✗ {fileName}: {kvp.Value}")
                                                 End If
                                             Next

                                             sb.AppendLine()
                                             sb.AppendLine($"Signed: {successCount} / {results.Count}")

                                             ShowCustomMessageBox(sb.ToString())
                                         End Sub

            lbFiles.BringToFront()
            form.ShowDialog()
        End Sub

#End Region

    End Class
End Namespace